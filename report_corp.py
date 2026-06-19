import io
import uuid
from pydoc import doc
import zipfile
import copy
import pandas as pd
import numpy as np
import docx
import os
import re 
from docxtpl import DocxTemplate
import matplotlib.pyplot as plt
from docxtpl import InlineImage
from docx.shared import Mm
import warnings
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from docx.oxml.ns import qn as _qn
from lxml import etree
import tempfile
import matplotlib.image as mpimg
import matplotlib.patches as patches

warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")


PERCORSO_LOGO_WATERMARK = "logofv.png" # WM

def aggiungi_watermark_fig(fig, applica):
    """Sovrappone logo e testo opachi, MANTENENDO intatte le dimensioni originali"""
    if not applica:
        return
        
    # Peschiamo l'asse originale del grafico per ancorarci a lui (evita sfasamenti e resize!)
    ax = fig.gca()
    
    # 1. 🛡️ VETRO SATINATO (alpha=0.96) ancorato all'area interna del grafico
    rect = patches.Rectangle((-0.1, -0.1), 1.2, 1.2, transform=ax.transAxes,
                             linewidth=0, edgecolor='none', facecolor='white', 
                             alpha=0.96, zorder=998, clip_on=False)
    ax.add_patch(rect)
    
    # 2. Creiamo un asse "fantasma" tenendolo rigorosamente DENTRO i margini (non allarga il grafico)
    ax_bg = fig.add_axes([0.1, 0.1, 0.8, 0.8], zorder=999)
    ax_bg.axis('off')
    
    try:
        # Inserisce il Logo a sinistra (Opacità 100%, Colori Vividi)
        logo = mpimg.imread(PERCORSO_LOGO_WATERMARK)
        ax_bg.imshow(logo, extent=[0.0, 0.40, 0.10, 0.90], alpha=1.0, aspect='auto')
        
        # Inserisce il Testo a destra (Opacità 100%, Colori Vividi)
        ax_bg.text(0.75, 0.5, "Finance & Value\nREPORT PREMIUM", fontsize=24, color='#2B3A67', 
                   alpha=1.0, ha='center', va='center', rotation=15, weight='bold')
    except:
        ax_bg.text(0.5, 0.5, "Finance & Value\nDATI BLOCCATI", fontsize=32, color='#2B3A67', 
                   alpha=1.0, ha='center', va='center', rotation=15, weight='bold')
        

def format_euro(numero, decimali=2):
    if pd.isna(numero): return "n.d."
    formato = f"{{:,.{decimali}f}}".format(numero)
    return formato.replace(',', 'X').replace('.', ',').replace('X', '.')


# =====================================================================
# ☢️ LA LAVATRICE NUCLEARE (Ricostruisce l'XML di Word da zero)
# =====================================================================

def lavatrice_nucleare(template_path):
    doc_temp = docx.Document(template_path)
    
    def ripara_paragrafo(p):
        testo = p.text
        # Se trova un tag di tabella o grafico
        if '{{' in testo and '}}' in testo and any(tag in testo for tag in ['{{ tabella_', '{{ grafico_', '{{tabella_', '{{grafico_']):
            # Taglia il testo usando i tag come forbici
            frammenti = re.split(r'(\{\{.*?\}\})', testo)
            
            for frammento in frammenti:
                pulito = frammento.strip()
                if pulito: # Se non è vuoto
                    # Crea un paragrafo puro e totalmente indipendente
                    nuovo_p = p.insert_paragraph_before('')
                    nuovo_p.clear()
                    nuovo_p.add_run(pulito)
                    try:
                        nuovo_p.style = p.style
                    except:
                        pass
            # Svuota il paragrafo originale corrotto, lasciando un a capo pulito obbligatorio
            p._p.getparent().remove(p._p)

    # 1. Scansiona tutto il documento
    for p in doc_temp.paragraphs:
        ripara_paragrafo(p)
        
    # 2. Scansiona le vecchie tabelle (se ci hai lasciato tag dentro per sbaglio)
    for table in doc_temp.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    ripara_paragrafo(p)
                # Obbliga ogni cella a finire con un a capo (Regola d'oro di Word)
                try:
                    cell.add_paragraph("")
                except:
                    pass

    temp_path = os.path.join(tempfile.gettempdir(), f"template_nucleare_{uuid.uuid4().hex[:8]}.docx")
    doc_temp.save(temp_path)
    return temp_path

# =====================================================================
# MOTORE DI CALCOLO E IMPAGINAZIONE
# =====================================================================
def genera_report_word(zip_buffer, template_path, azienda_target, df_orbis, settore_nace, num_max_soc_orbis, modalita_teaser=False):
    
    # 🛡️ ANTIDOTO: Disintegra i caratteri invisibili di Excel che corrompono Word
    caratteri_proibiti = dict.fromkeys(range(0, 32))
    caratteri_proibiti.pop(9, None); caratteri_proibiti.pop(10, None); caratteri_proibiti.pop(13, None)
    azienda_target = str(azienda_target).translate(caratteri_proibiti)
    settore_nace = str(settore_nace).translate(caratteri_proibiti)
    df_orbis = df_orbis.replace(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', regex=True)

    df_orbis = df_orbis.copy()

    # --- PULIZIA DATI INIZIALE (Elimina tutti gli n.d. testuali) ---
    df_orbis = df_orbis.replace(['n.d.', 'n.a.', 'n.s.', 'N.D.', 'N.A.', 'N.S.', 'n.d', ' '], np.nan)
    # ---------------------------------------------------------------

    colonne_da_convertire = [
        'Totale valore della produzione migl EUR 2024', 'Totale Attivo migl EUR 2024',
        'Numero dipendenti 2024', 'Margine di Profitto (*) % 2024', 'Margine EBITDA (*) % 2024',
        'Margine EBIT (*) % 2024', 'Indice di Struttura 1° livello (*) 2024',
        'Indice di Struttura 2° livello (*) 2024', 'Gearing (*) % 2024',
        'Indice di Rotazione del Capitale Investito (*) 2024', 'Current Ratio (*) 2024', 'Quick Ratio (*) 2024'
    ]
    for c in colonne_da_convertire:
        if c in df_orbis.columns:
            df_orbis[c] = pd.to_numeric(df_orbis[c], errors='coerce')

    col_ragione = [c for c in df_orbis.columns if 'ragione' in str(c).lower()][0]
    col_nuts = [c for c in df_orbis.columns if 'nuts2' in str(c).lower() or 'nuts 2' in str(c).lower()]
    col_regione = col_nuts[0] if col_nuts else None

    df_orbis['Forma Giuridica Pulita'] = df_orbis['Forma giuridica nazionale'].astype(str).str.replace(r'\s*\(Italia\)', '', regex=True).str.strip()
    
    # 🔠 DIZIONARIO ABBREVIAZIONI FORME GIURIDICHE
    mappatura_fg = {
        'società a responsabilità limitata semplificata': 'S.r.l.s.',
        "societa' a responsabilita' limitata semplificata": 'S.r.l.s.',
        "societa a responsabilita limitata semplificata": 'S.r.l.s.',
        'società a responsabilità limitata': 'S.r.l.',
        "societa' a responsabilita' limitata": 'S.r.l.',
        "societa a responsabilita limitata": 'S.r.l.',
        'società per azioni': 'S.p.A.',
        "societa' per azioni": 'S.p.A.',
        "societa per azioni": 'S.p.A.',
        'società in nome collettivo': 'S.n.c.',
        "societa' in nome collettivo": 'S.n.c.',
        "societa in nome collettivo": 'S.n.c.',
        'società in accomandita semplice': 'S.a.s.',
        "societa' in accomandita semplice": 'S.a.s.',
        "societa in accomandita semplice": 'S.a.s.',
        'società in accomandita per azioni': 'S.a.p.a.',
        "societa' in accomandita per azioni": 'S.a.p.a.',
        "societa in accomandita per azioni": 'S.a.p.a.',
        'società cooperativa': 'Soc. Coop.',
        "societa' cooperativa": 'Soc. Coop.',
        "societa cooperativa": 'Soc. Coop.',
        'società semplice': 'S.s.',
        "societa' semplice": 'S.s.',
        "societa semplice": 'S.s.',
        'società consortile': 'Soc. Cons.',
        "societa' consortile": 'Soc. Cons.',
        "societa consortile": 'Soc. Cons.'
    }

    def abbrevia_forma_giuridica(fg_estesa):
        if pd.isna(fg_estesa): return "N.D."
        fg_lower = str(fg_estesa).lower().strip()
        for chiave, abbreviazione in mappatura_fg.items():
            if chiave in fg_lower: return abbreviazione
        if "s.r.l." in fg_lower or "srl" in fg_lower: return "S.r.l."
        if "s.p.a." in fg_lower or "spa" in fg_lower: return "S.p.A."
        return str(fg_estesa).title()

    df_orbis['Forma Giuridica Pulita'] = df_orbis['Forma Giuridica Pulita'].apply(abbrevia_forma_giuridica)

    def get_macro(nuts2):
        code = str(nuts2)[:3]
        if code == 'ITC': return 'Nord Ovest'
        elif code == 'ITH': return 'Nord Est'
        elif code == 'ITI': return 'Centro'
        elif code in ['ITF', 'ITG']: return 'Sud e Isole'
        else: return 'Altro'
    
    if col_regione: df_orbis['Macroregione'] = df_orbis[col_regione].apply(get_macro)
    else: df_orbis['Macroregione'] = 'Altro'

    tot_imprese_settore = len(df_orbis)
    tot_ricavi_settore = df_orbis['Totale valore della produzione migl EUR 2024'].sum()
    tot_attivo_settore = df_orbis['Totale Attivo migl EUR 2024'].sum()
    tot_dipendenti_settore = df_orbis['Numero dipendenti 2024'].sum()
    
    fg_counts = df_orbis['Forma Giuridica Pulita'].value_counts()
    fg_maggioranza = fg_counts.index[0] if not fg_counts.empty else "n.d."
    num_fg_maggioranza = fg_counts.iloc[0] if not fg_counts.empty else 0
    perc_fg_maggioranza = (num_fg_maggioranza / tot_imprese_settore) * 100 if tot_imprese_settore > 0 else 0

    df_target = df_orbis[df_orbis[col_ragione].astype(str).str.lower().str.contains(azienda_target.lower().strip(), regex=False, na=False)]
    
    def get_val_and_rank(col_name, is_lower_better=False):
        if col_name not in df_orbis.columns or df_target.empty: return "n.d.", "n.d.", "n.d."
        target_idx = df_target.index[0]
        valore = df_target.at[target_idx, col_name]
        rank_naz = df_orbis[col_name].rank(ascending=is_lower_better, method='min')
        pos_naz = int(rank_naz.loc[target_idx]) if pd.notna(rank_naz.loc[target_idx]) else "n.d."
        if col_regione:
            rank_reg = df_orbis.groupby(col_regione)[col_name].rank(ascending=is_lower_better, method='min')
            pos_reg = int(rank_reg.loc[target_idx]) if pd.notna(rank_reg.loc[target_idx]) else "n.d."
        else: pos_reg = "n.d."
        return format_euro(valore), pos_naz, pos_reg

    mg_prof, rnk_naz_prof, rnk_reg_prof = get_val_and_rank('Margine di Profitto (*) % 2024', False)
    mg_ebitda, rnk_naz_ebitda, rnk_reg_ebitda = get_val_and_rank('Margine EBITDA (*) % 2024', False)
    mg_ebit, rnk_naz_ebit, rnk_reg_ebit = get_val_and_rank('Margine EBIT (*) % 2024', False)
    ind_str1, rnk_naz_strut1, rnk_reg_strut1 = get_val_and_rank('Indice di Struttura 1° livello (*) 2024', False)
    ind_str2, rnk_naz_strut2, rnk_reg_strut2 = get_val_and_rank('Indice di Struttura 2° livello (*) 2024', False)
    gearing, rnk_naz_gear, rnk_reg_gear = get_val_and_rank('Gearing (*) % 2024', True) 
    ind_rot, rnk_naz_rot, rnk_reg_rot = get_val_and_rank('Indice di Rotazione del Capitale Investito (*) 2024', False)
    ind_cr, rnk_naz_cr, rnk_reg_cr = get_val_and_rank('Current Ratio (*) 2024', False)
    ind_qr, rnk_naz_qr, rnk_reg_qr = get_val_and_rank('Quick Ratio (*) 2024', False)

    if not df_target.empty:
        riga = df_target.iloc[0]
        p_iva = riga.get('Codice fiscale/Partita IVA', 'N.D.')
        forma_giuridica = riga.get('Forma Giuridica Pulita', 'N.D.')
        macroregione_target = riga.get('Macroregione', 'N.D.')
        num_fg_target = fg_counts.get(forma_giuridica, 0)
        perc_fg_target = (num_fg_target / tot_imprese_settore) * 100 if tot_imprese_settore > 0 else 0
        ricavi_mgl = riga.get('Totale valore della produzione migl EUR 2024')
        attivo_mgl = riga.get('Totale Attivo migl EUR 2024')
        dipendenti = riga.get('Numero dipendenti 2024')
        ricavi_mln = ricavi_mgl / 1000 if pd.notna(ricavi_mgl) else 0
        attivo_mln = attivo_mgl / 1000 if pd.notna(attivo_mgl) else 0
        perc_ricavi_panel = (ricavi_mgl / tot_ricavi_settore * 100) if tot_ricavi_settore > 0 and pd.notna(ricavi_mgl) else 0
        perc_attivo_panel = (attivo_mgl / tot_attivo_settore * 100) if tot_attivo_settore > 0 and pd.notna(attivo_mgl) else 0
        tot_dip_area = df_orbis[df_orbis['Macroregione'] == macroregione_target]['Numero dipendenti 2024'].sum()
        perc_dip_area = (dipendenti / tot_dip_area * 100) if tot_dip_area > 0 and pd.notna(dipendenti) else 0
        # --- NUOVI CALCOLI TERRITORIALI ---
        regione_grezza = str(riga.get(col_regione, 'N.D.'))
        regione_target_pulita = regione_grezza.split(' - ')[1] if ' - ' in regione_grezza else regione_grezza
        tot_imprese_regione = len(df_orbis[df_orbis[col_regione] == riga.get(col_regione)]) if col_regione else 0
        perc_imprese_regione = (tot_imprese_regione / tot_imprese_settore * 100) if tot_imprese_settore > 0 else 0
        tot_ricavi_macro_mgl = df_orbis[df_orbis['Macroregione'] == macroregione_target]['Totale valore della produzione migl EUR 2024'].sum()
        tot_ricavi_macro_mln = tot_ricavi_macro_mgl / 1000
        perc_ricavi_macroregione = (tot_ricavi_macro_mgl / tot_ricavi_settore * 100) if tot_ricavi_settore > 0 else 0
        perc_ricavi_target_su_macro = (ricavi_mgl / tot_ricavi_macro_mgl * 100) if tot_ricavi_macro_mgl > 0 and pd.notna(ricavi_mgl) else 0

        # =======================================================
        # --- NUOVI CALCOLI CATEGORIA (PER FORMA GIURIDICA) ---
        # =======================================================
        # Filtriamo il database prendendo SOLO le aziende con la stessa forma giuridica (es. solo le S.p.A.)
        df_categoria = df_orbis[df_orbis['Forma Giuridica Pulita'] == forma_giuridica]
        
        tot_ricavi_categoria = df_categoria['Totale valore della produzione migl EUR 2024'].sum()
        tot_attivo_categoria = df_categoria['Totale Attivo migl EUR 2024'].sum()
        
        perc_ricavi_categoria = (ricavi_mgl / tot_ricavi_categoria * 100) if tot_ricavi_categoria > 0 and pd.notna(ricavi_mgl) else 0
        perc_attivo_categoria = (attivo_mgl / tot_attivo_categoria * 100) if tot_attivo_categoria > 0 and pd.notna(attivo_mgl) else 0

        try: quartile_ricavi_target = pd.qcut(df_orbis['Totale valore della produzione migl EUR 2024'].dropna(), 4, labels=[1, 2, 3, 4]).loc[riga.name]
        except: quartile_ricavi_target = "N.D."
        try: quartile_attivo_target = pd.qcut(df_orbis['Totale Attivo migl EUR 2024'].dropna(), 4, labels=[1, 2, 3, 4]).loc[riga.name]
        except: quartile_attivo_target = "N.D."
    else:
        p_iva, forma_giuridica, macroregione_target, num_fg_target, perc_fg_target, ricavi_mln, attivo_mln, dipendenti = "N.D.", "N.D.", "N.D.", 0, 0, 0, 0, 0
        perc_ricavi_panel, perc_attivo_panel, perc_dip_area, quartile_ricavi_target, quartile_attivo_target = 0, 0, 0, "N.D.", "N.D."
        regione_target_pulita = "N.D."
        perc_imprese_regione, tot_ricavi_macro_mln, perc_ricavi_macroregione, perc_ricavi_target_su_macro = 0, 0, 0, 0
        perc_ricavi_categoria, perc_attivo_categoria = 0, 0  # <--- AGGIUNTA QUESTA RIGA!

    # 🧽 PULITURA RAGIONE SOCIALE
    ragione_sociale_pulita = str(azienda_target)
    sostituzioni_nomi = {
        "SOCIETA' A RESPONSABILITA' LIMITATA SEMPLIFICATA": "S.r.l.s.",
        "SOCIETÀ A RESPONSABILITÀ LIMITATA SEMPLIFICATA": "S.r.l.s.",
        "SOCIETA A RESPONSABILITA LIMITATA SEMPLIFICATA": "S.r.l.s.",
        "SOCIETA' A RESPONSABILITA' LIMITATA": "S.r.l.",
        "SOCIETÀ A RESPONSABILITÀ LIMITATA": "S.r.l.",
        "SOCIETA A RESPONSABILITA LIMITATA": "S.r.l.",
        "SOCIETA' PER AZIONI": "S.p.A.",
        "SOCIETÀ PER AZIONI": "S.p.A.",
        "SOCIETA PER AZIONI": "S.p.A.",
        "SOCIETA' IN NOME COLLETTIVO": "S.n.c.",
        "SOCIETÀ IN NOME COLLETTIVO": "S.n.c.",
        "SOCIETA IN NOME COLLETTIVO": "S.n.c.",
        "SOCIETA' IN ACCOMANDITA SEMPLICE": "S.a.s.",
        "SOCIETÀ IN ACCOMANDITA SEMPLICE": "S.a.s.",
        "SOCIETA IN ACCOMANDITA SEMPLICE": "S.a.s.",
        "SOCIETA' IN ACCOMANDITA PER AZIONI": "S.a.p.a.",
        "SOCIETÀ IN ACCOMANDITA PER AZIONI": "S.a.p.a.",
        "SOCIETA IN ACCOMANDITA PER AZIONI": "S.a.p.a.",
        "SOCIETA' COOPERATIVA": "Soc. Coop.",
        "SOCIETÀ COOPERATIVA": "Soc. Coop.",
        "SOCIETA COOPERATIVA": "Soc. Coop.",
        "SOCIETA' SEMPLICE": "S.s.",
        "SOCIETÀ SEMPLICE": "S.s.",
        "SOCIETA SEMPLICE": "S.s.",
        "SOCIETA' CONSORTILE": "Soc. Cons.",
        "SOCIETÀ CONSORTILE": "Soc. Cons.",
        "SOCIETA CONSORTILE": "Soc. Cons."
    }
    for lungo, corto in sostituzioni_nomi.items():
        ragione_sociale_pulita = re.sub(re.escape(lungo), corto, ragione_sociale_pulita, flags=re.IGNORECASE)

    # 🏷️ PULITURA NACE MULTIPLI
    settore_nace_str = str(settore_nace).strip()
    codici_estratti = []
    descrizioni_estratte = []
    matches = list(re.finditer(r'(\d{3,4})\s*-\s*(.*?)(?=\s*(?:[,;|]|\n)?\s*\d{3,4}\s*-|$)', settore_nace_str))

    if matches:
        for match in matches:
            cod_grezzo = match.group(1).strip()
            desc_testo = match.group(2).strip().rstrip(",;| ")
            cod_pulito = f"{cod_grezzo[:2]}.{cod_grezzo[2:]}"
            codici_estratti.append(cod_pulito)
            descrizioni_estratte.append(desc_testo)
            
        def formatta_italiano(lista):
            if len(lista) == 1: return lista[0]
            return ", ".join(lista[:-1]) + " e " + lista[-1]
            
        cod_nace_pulito = formatta_italiano(codici_estratti)
        desc_nace_pulita = formatta_italiano(descrizioni_estratte)
    else:
        cod_nace_pulito = "N.D."
        desc_nace_pulita = settore_nace_str

    # 🏛️ ESPANSIONE FORMA GIURIDICA (Per il testo discorsivo)
    dizionario_espansioni = {
        'S.r.l.s.': 'Società a Responsabilità Limitata Semplificata',
        'S.r.l.': 'Società a Responsabilità Limitata',
        'S.p.A.': 'Società per Azioni',
        'S.n.c.': 'Società in Nome Collettivo',
        'S.a.s.': 'Società in Accomandita Semplice',
        'S.a.p.a.': 'Società in Accomandita per Azioni',
        'Soc. Coop.': 'Società Cooperativa',
        'S.s.': 'Società Semplice',
        'Soc. Cons.': 'Società Consortile'
    }
    fg_espansa = dizionario_espansioni.get(forma_giuridica, forma_giuridica)


    # 🟢 FIX P.IVA: Ripristina gli zeri iniziali "mangiati" da Excel/Pandas
    if pd.notna(p_iva) and str(p_iva).strip().lower() not in ['n.d.', 'nan', '']:
        try:
            # 1. Lo passiamo a float e poi a int per eliminare eventuali ".0" finali
            p_iva_pulita = str(int(float(p_iva)))
            # 2. Aggiunge gli zeri a sinistra finché la lunghezza non torna a 11 cifre
            p_iva = p_iva_pulita.zfill(11)
        except ValueError:
            # Se dovesse contenere lettere (es. codici esteri o errori), lo lascia a testo
            p_iva = str(p_iva)
    else:
        p_iva = "n.d."


    # COSTRUZIONE DEL DIZIONARIO (Dati per il Word)
    context = {
        'ragione_sociale': ragione_sociale_pulita, 
        'codice_nace': cod_nace_pulito, 
        'descr_settore': desc_nace_pulita, 
        'partita_iva': p_iva, 
        'forma_giuridica': forma_giuridica,
        'forma_giuridica_espansa': fg_espansa,
        'perc_fg': format_euro(perc_fg_target), 'num_fg': f"{num_fg_target:,}".replace(',', '.'),
        'fg_maggioranza': fg_maggioranza, 'num_fg_maggioranza': f"{num_fg_maggioranza:,}".replace(',', '.'),
        'perc_fg_maggioranza': format_euro(perc_fg_maggioranza),
        'classe_dimensionale': 'Grande Impresa' if ricavi_mln > 50 else ('Media Impresa' if ricavi_mln > 10 else 'Piccola Impresa'),
        'ricavi_mln': format_euro(ricavi_mln),
        'perc_ricavi_panel': format_euro(perc_ricavi_panel),
        'perc_ricavi_categoria': format_euro(perc_ricavi_categoria), # <-- Ora calcola il VERO dato!
        'quartile_ricavi': quartile_ricavi_target,
        'attivo_mln': format_euro(attivo_mln),
        'perc_attivo_panel': format_euro(perc_attivo_panel),
        'perc_attivo_categoria': format_euro(perc_attivo_categoria), # <-- Aggiunto anche per l'Attivo!
        'quartile_attivo': quartile_attivo_target,
        'perc_dip_area': format_euro(perc_dip_area), 'macroregione': macroregione_target,
        'regione_target': regione_target_pulita,
        'perc_imprese_regione': format_euro(perc_imprese_regione),
        'perc_ricavi_macroregione': format_euro(perc_ricavi_macroregione),
        'tot_ricavi_macro_mln': format_euro(tot_ricavi_macro_mln),
        'perc_ricavi_target_su_macro': format_euro(perc_ricavi_target_su_macro),
        'mg_ebitda': mg_ebitda, 'rnk_naz_ebitda': rnk_naz_ebitda, 'rnk_reg_ebitda': rnk_reg_ebitda,
        'mg_ebit': mg_ebit, 'rnk_naz_ebit': rnk_naz_ebit, 'rnk_reg_ebit': rnk_reg_ebit,
        'mg_prof': mg_prof, 'rnk_naz_prof': rnk_naz_prof, 'rnk_reg_prof': rnk_reg_prof,
        'ind_str1': ind_str1, 'rnk_naz_strut1': rnk_naz_strut1, 'rnk_reg_strut1': rnk_reg_strut1,
        'ind_str2': ind_str2, 'rnk_naz_strut2': rnk_naz_strut2, 'rnk_reg_strut2': rnk_reg_strut2,
        'gearing': gearing, 'rnk_naz_gear': rnk_naz_gear, 'rnk_reg_gear': rnk_reg_gear,
        'ind_rot_cap': ind_rot, 'rnk_naz_rot': rnk_naz_rot, 'rnk_reg_rot': rnk_reg_rot,
        'ind_cr': ind_cr, 'rnk_naz_cr': rnk_naz_cr, 'rnk_reg_cr': rnk_reg_cr,
        'ind_qr': ind_qr, 'rnk_naz_qr': rnk_naz_qr, 'rnk_reg_qr': rnk_reg_qr,
        'tot_imprese': f"{tot_imprese_settore:,}".replace(',', '.'), 'tot_ricavi': format_euro(tot_ricavi_settore),
        'tot_attivo': format_euro(tot_attivo_settore), 'tot_dipendenti': format_euro(tot_dipendenti_settore, 0),
        
        'perc_ricavi_target_su_macro': format_euro(perc_ricavi_target_su_macro),
        'impatto_territoriale': 'N.D.',

        # Variabili di default (sovrascritte dal motore sottostante)
        'rating_eco': 'N.D.', 'rating_patr': 'N.D.', 'rating_fin': 'N.D.', 'rating_tot': 'N.D.', 'rating_comb': 'N.D.',
        'rating_piu_presente': 'N.D.', 'assoc_desc_rating_lett': 'N.D.',
        'rat1_piu_pres_num': 'N.D.', 'rat1_piu_pres_categ': 'N.D.', 
        'rat2_piu_pres_num': 'N.D.', 'rat2_piu_pres_categ': 'N.D.',
        'rat3_piu_pres_num': 'N.D.', 'rat3_piu_pres_categ': 'N.D.', 
        'rating_piu_pres': 'N.D.', 'rating_piu_pres_num_tot': 'N.D.',
        'num_max_soc': 'N.D.', 'num_soc_valide': 'N.D.', 'perc_su_istat': '100', 'max_soc_istat': 'N.D.',
        'tab_territorio': [], 'tab_bench_territorio': []
    }

    # =================================================================
    # 📊 MOTORE DI CALCOLO NATIVO DEI RATING (Senza Excel!)
    # =================================================================
    def punteggio_diretto(val, t1, t2):
        if pd.isna(val): return 1
        return 3 if val >= t2 else (2 if val >= t1 else 1)

    def punteggio_inverso(val, t1, t2):
        if pd.isna(val): return 1
        return 3 if val <= t1 else (2 if val <= t2 else 1)

    def assegna_lettera(punti):
        if pd.isna(punti): return 'C'
        return 'A' if punti >= 8 else ('B' if punti >= 5 else 'C')

    # Identificatori metriche
    c_prof = 'Margine di Profitto (*) % 2024'
    c_ebitda = 'Margine EBITDA (*) % 2024'
    c_ebit = 'Margine EBIT (*) % 2024'
    c_rot = 'Indice di Rotazione del Capitale Investito (*) 2024'
    c_quick = 'Quick Ratio (*) 2024'
    c_curr = 'Current Ratio (*) 2024'
    c_str1 = 'Indice di Struttura 1° livello (*) 2024'
    c_str2 = 'Indice di Struttura 2° livello (*) 2024'
    c_gear = 'Gearing (*) % 2024'

    metriche_dirette = [c_prof, c_ebitda, c_ebit, c_rot, c_quick, c_curr, c_str1, c_str2]
    metriche_inverse = [c_gear]

    df_rating = df_orbis.copy()

    # 1. Calcolo matematico dei Terzili e Assegnazione Punti
    for m in metriche_dirette + metriche_inverse:
        if m in df_rating.columns:
            t1 = df_rating[m].quantile(1/3)
            t2 = df_rating[m].quantile(2/3)
            if m in metriche_inverse:
                df_rating[f'pts_{m}'] = df_rating[m].apply(lambda x: punteggio_inverso(x, t1, t2))
            else:
                df_rating[f'pts_{m}'] = df_rating[m].apply(lambda x: punteggio_diretto(x, t1, t2))
        else:
            df_rating[f'pts_{m}'] = 1  # Fallback se colonna non esiste

    # 2. Somma Punti per Area (Economico, Finanziario, Patrimoniale)
    df_rating['pts_eco'] = df_rating[f'pts_{c_prof}'] + df_rating[f'pts_{c_ebitda}'] + df_rating[f'pts_{c_ebit}']
    df_rating['pts_fin'] = df_rating[f'pts_{c_rot}'] + df_rating[f'pts_{c_quick}'] + df_rating[f'pts_{c_curr}']
    df_rating['pts_pat'] = df_rating[f'pts_{c_str1}'] + df_rating[f'pts_{c_str2}'] + df_rating[f'pts_{c_gear}']

    # 3. Assegnazione Lettere (A, B, C)
    df_rating['Rating Economico'] = df_rating['pts_eco'].apply(assegna_lettera)
    df_rating['Rating Finanziario'] = df_rating['pts_fin'].apply(assegna_lettera)
    df_rating['Rating Patrimoniale'] = df_rating['pts_pat'].apply(assegna_lettera)

    # 4. Rating Combinato Testuale (es: AAA, BBB, ABC)
    df_rating['Rating Combinato'] = df_rating['Rating Economico'] + df_rating['Rating Patrimoniale'] + df_rating['Rating Finanziario']

    # 5. Benchmark Totale (Media delle Lettere)
    valori_lettere = {'A': 3, 'B': 2, 'C': 1}
    df_rating['pts_totali'] = df_rating['Rating Economico'].map(valori_lettere) + \
                              df_rating['Rating Finanziario'].map(valori_lettere) + \
                              df_rating['Rating Patrimoniale'].map(valori_lettere)
    df_rating['Benchmark Totale'] = df_rating['pts_totali'].apply(assegna_lettera)

    
    # =================================================================
    # 🎯 ESTRAZIONE E INIEZIONE NEL TEMPLATE WORD
    # =================================================================
    
    # Preleviamo la riga esatta dell'azienda target
    df_target_rating = df_rating[df_rating[col_ragione].astype(str).str.lower().str.contains(azienda_target.lower().strip(), regex=False, na=False)]

    if not df_target_rating.empty:
        riga_t = df_target_rating.iloc[0]
        context['rating_eco'] = riga_t['Rating Economico']
        context['rating_fin'] = riga_t['Rating Finanziario']
        context['rating_patr'] = riga_t['Rating Patrimoniale']
        context['rating_tot'] = riga_t['Rating Combinato']
        context['rating_comb'] = riga_t['Benchmark Totale']

        context['dip_target'] = f"{int(dipendenti):,}".replace(',', '.') if pd.notna(dipendenti) else "n.d."
       
        # Calcoli matematici reali per il Box Economico
        r_eco_target = riga_t['Rating Economico']
        num_eco_fascia_real = len(df_rating[df_rating['Rating Economico'] == r_eco_target])
        perc_eco_fascia_real = (num_eco_fascia_real / tot_imprese_settore) * 100 if tot_imprese_settore > 0 else 0
       
        context['nr_rating_eco'] = f"{num_eco_fascia_real:,}".replace(',', '.')
        context['perc_rating_eco'] = format_euro(perc_eco_fascia_real)
        context['perc_rating_eco_parte'] = format_euro((1 / num_eco_fascia_real) * 100) if num_eco_fascia_real > 0 else "0,00"
       
        # Calcoli matematici reali per il Box Patrimoniale
        r_patr_target = riga_t['Rating Patrimoniale']
        num_patr_fascia_real = len(df_rating[df_rating['Rating Patrimoniale'] == r_patr_target])
        perc_patr_fascia_real = (num_patr_fascia_real / tot_imprese_settore) * 100 if tot_imprese_settore > 0 else 0
       
        context['nr_rating_patr'] = f"{num_patr_fascia_real:,}".replace(',', '.')
        context['perc_rating_patr'] = format_euro(perc_patr_fascia_real)
        context['perc_rating_patr_parte'] = format_euro((1 / num_patr_fascia_real) * 100) if num_patr_fascia_real > 0 else "0,00"

        # Calcoli matematici reali per il Box Finanziario
        r_fin_target = riga_t['Rating Finanziario']
        num_fin_fascia_real = len(df_rating[df_rating['Rating Finanziario'] == r_fin_target])
        perc_fin_fascia_real = (num_fin_fascia_real / tot_imprese_settore) * 100 if tot_imprese_settore > 0 else 0
       
        context['nr_rating_fin'] = f"{num_fin_fascia_real:,}".replace(',', '.')
        context['perc_rating_fin'] = format_euro(perc_fin_fascia_real)
        context['perc_rating_fin_parte'] = format_euro((1 / num_fin_fascia_real) * 100) if num_fin_fascia_real > 0 else "0,00"


    # --- 🤖 MOTORE NARRATIVO (Testi Dinamici in base al Rating) ---

    def get_impatto_territoriale(perc, nome, ricavi_formattati):
        if perc >= 5.0:
            return f"Con ricavi pari a {ricavi_formattati} mln di EUR, {nome} incide in maniera determinante sulla creazione di ricchezza locale, confermandosi un player di assoluto riferimento sul piano territoriale grazie a un impatto del {format_euro(perc)}% rispetto al totale dei ricavi dell'area."
        elif perc >= 1.0:
            return f"Con ricavi pari a {ricavi_formattati} mln di EUR, {nome} fornisce un contributo significativo alla creazione di ricchezza locale, consolidando una posizione di rilievo sul piano territoriale con un'incidenza del {format_euro(perc)}% rispetto ai ricavi complessivi dell'area."
        elif perc >= 0.1:
            return f"Con ricavi pari a {ricavi_formattati} mln di EUR, {nome} partecipa attivamente al tessuto economico locale, rappresentando una stabile e sana realtà territoriale con un'incidenza del {format_euro(perc)}% rispetto ai ricavi complessivi dell'area."
        else:
            return f"Con ricavi pari a {ricavi_formattati} mln di EUR, {nome} opera all'interno di un mercato territoriale ampio e competitivo, contribuendo in modo fisiologico al tessuto economico locale con un'incidenza del {format_euro(perc)}% rispetto ai ricavi complessivi dell'area."
    
    def get_testo_totale(rating):
        if rating == 'A': return "Questo risultato riflette un profilo di eccellenza, stabilità e robustezza strutturale, posizionando l'azienda ai vertici del settore."
        elif rating == 'B': return "Questo risultato riflette un profilo di sostanziale stabilità, evidenziando una struttura solida pur con alcuni margini di miglioramento in ambiti specifici."
        elif rating == 'C': return "Questo risultato evidenzia elementi di vulnerabilità complessiva, suggerendo la necessità di interventi mirati per stabilizzare e rinforzare la struttura aziendale."
        return "Non sono disponibili dati sufficienti per elaborare un giudizio complessivo accurato."

    def get_testo_eco(rating):
        if rating == 'A': return "La redditività si presenta eccellente e superiore alla media di settore, dimostrando una notevole efficienza nel trasformare il valore della produzione in profitto e un'ottima gestione dei costi operativi."
        elif rating == 'B': return "La redditività risulta adeguata e in linea con le potenzialità dell'azienda, seppur con spazi di ottimizzazione. Un'attenta razionalizzazione dei costi potrebbe incrementare ulteriormente la trasformazione del valore in profitto netto."
        elif rating == 'C': return "La redditività si presenta inferiore alle potenzialità dell'azienda, penalizzata da un elevato peso dei costi di gestione. I margini operativi si collocano al di sotto della media settoriale, segnalando l'urgenza di adottare misure mirate alla razionalizzazione dei costi."
        return "Dati economici non disponibili o insufficienti."

    def get_testo_patr(rating):
        if rating == 'A': return "Costituisce il punto di forza dell'azienda, che si distingue per l'assenza di rischi di insolvenza a lungo termine. Il Patrimonio Netto copre ampiamente l'attivo fisso, assicurando una solidità ineccepibile e un'autonomia economica di assoluto rilievo."
        elif rating == 'B': return "La struttura patrimoniale appare equilibrata. L'azienda mostra una discreta indipendenza finanziaria, sebbene il ricorso al capitale di terzi debba essere monitorato per mantenere intatta la stabilità nel lungo periodo."
        elif rating == 'C': return "La struttura evidenzia elementi di debolezza, con un livello di indebitamento potenzialmente critico rispetto al capitale proprio. È consigliabile un ribilanciamento delle fonti di finanziamento per mitigare i rischi a lungo termine."
        return "Dati patrimoniali non disponibili o insufficienti."

    def get_testo_fin(rating):
        if rating == 'A': return "La gestione della liquidità è ottimale. L'azienda genera ampi flussi di cassa ed evidenzia un'eccellente solvibilità a breve termine, garantendo la totale flessibilità nell'affrontare impegni correnti e contingenze."
        elif rating == 'B': return "Si riscontra una solvibilità a breve termine sufficiente, pur con una flessibilità moderata nella gestione della liquidità. Le risorse coprono gli impegni, ma si suggerisce un monitoraggio costante del capitale circolante."
        elif rating == 'C': return "Nonostante la continuità operativa, si riscontra una scarsa flessibilità nella gestione della liquidità immediata. Una quota rilevante delle risorse risulta vincolata in attività correnti difficilmente trasformabili in denaro liquido, esponendo l'azienda a potenziali difficoltà."
        return "Dati finanziari non disponibili o insufficienti."

    def get_testo_sintesi(rating):
        if rating == 'A': return "In sintesi, l’azienda vanta fondamenta estremamente solide, garantendo continuità aziendale e ottime prospettive di sviluppo. Mantenendo questo approccio strategico virtuoso, le risorse potranno essere focalizzate serenamente su investimenti di espansione aziendale."
        elif rating == 'B': return "In sintesi, l'azienda si poggia su basi strutturali solide per garantire la continuità operativa. Per migliorare la propria competitività, sarà cruciale concentrare gli sforzi strategici su una gestione ancora più efficace delle aree meno performanti per liberare ulteriori risorse utili alla crescita."
        elif rating == 'C': return "In sintesi, l'azienda affronta sfide strutturali per garantire una serena continuità aziendale nel lungo termine. Per ripristinare la competitività, sarà cruciale concentrare gli sforzi strategici su una gestione drastica della riduzione dei costi e un attento recupero di liquidità immediata."
        return ""

    def get_intro_benchmark_eco(rating):
        if rating == 'A': return f'Rispetto al Benchmark Economico, la valutazione "{rating}" sottolinea una redditività operativa e netta di assoluta eccellenza, nettamente superiore alle performance medie del settore di riferimento, come visibile di seguito:'
        elif rating == 'B': return f'Rispetto al Benchmark Economico, la valutazione "{rating}" evidenzia una redditività operativa e netta adeguata e in linea con le medie del settore di riferimento, pur con fisiologici margini di miglioramento, come visibile di seguito:'
        elif rating == 'C': return f'Rispetto al Benchmark Economico, la valutazione "{rating}" sottolinea una redditività operativa e netta al di sotto delle potenzialità medie del settore di riferimento, come visibile di seguito:'
        return "Rispetto al Benchmark Economico, i dati a disposizione non consentono di esprimere una valutazione completa, come riassunto di seguito:"

    def get_analisi_margini_operativi(rating):
        if rating == 'A': return "evidenziano un'eccellente capacità di convertire i ricavi in margini operativi. Tali valori indicano un'ottima ottimizzazione dei costi di gestione e un'elevata efficienza caratteristica rispetto al settore di riferimento;"
        elif rating == 'B': return "evidenziano una buona capacità di convertire i ricavi in margini operativi. I valori risultano adeguati al settore di riferimento, pur lasciando spazio a un'ulteriore e fisiologica razionalizzazione dei costi di gestione;"
        elif rating == 'C': return "evidenziano una limitata capacità di convertire i ricavi in margini operativi. Tali valori indicano una significativa incidenza dei costi di gestione rispetto agli standard del settore di riferimento, suggerendo la necessità di recuperare efficienza;"
        return "non consentono una valutazione accurata della capacità di conversione dei ricavi in margini operativi."

    def get_analisi_margine_profitto(rating):
        if rating == 'A': return "conferma un'elevata efficienza complessiva. L'azienda non solo genera ricavi, ma riesce a trattenerne una percentuale considerevole come utile netto, a dimostrazione di una gestione fortemente redditizia."
        elif rating == 'B': return "segnala una redditività stabile. L'azienda genera ricavi trattenendo una percentuale equilibrata come utile netto, in linea con le dinamiche medie della catena del valore settoriale."
        elif rating == 'C': return "conferma la capacità di generare ricavi, ma segnala al contempo una bassa efficienza complessiva nella catena del valore, trattenendo solo una modesta percentuale come utile netto."
        return "non fornisce elementi sufficienti per un'analisi dettagliata della redditività netta."
    
    def get_intro_benchmark_patr(rating):
        if rating == 'A': return "attesta una straordinaria solidità strutturale dell'azienda, evidenziandone un'ottima copertura dai rischi a lungo termine. Questo risultato la colloca come un'eccellenza rispetto alla media del settore di riferimento:"
        elif rating == 'B': return "evidenzia una solidità strutturale adeguata, con un buon livello di copertura dai rischi a lungo termine. Questo risultato riflette una situazione di generale equilibrio rispetto alla media del settore di riferimento:"
        elif rating == 'C': return "segnala elementi di vulnerabilità strutturale, evidenziando una potenziale esposizione ai rischi a lungo termine. Questo risultato colloca l'azienda al di sotto delle performance medie del settore di riferimento, richiedendo attenzione su:"
        return "non consente di esprimere una valutazione completa sui rischi a lungo termine a causa di dati insufficienti:"

    def get_analisi_indici_struttura(rating):
        if rating == 'A': return "confermano che il Patrimonio Netto copre abbondantemente l'intero attivo fisso. In un contesto in cui gli investimenti strutturali richiedono stabilità, una tale copertura con risorse proprie rappresenta un chiaro indicatore di massima sicurezza e robustezza."
        elif rating == 'B': return "mostrano valori che indicano una discreta copertura dell'attivo fisso tramite il Patrimonio Netto. Pur garantendo la continuità operativa, vi sono i margini per consolidare ulteriormente la robustezza patrimoniale nel medio-lungo periodo."
        elif rating == 'C': return "evidenziano valori che indicano una parziale o insufficiente copertura dell'attivo fisso tramite mezzi propri. Questa situazione suggerisce un'eccessiva dipendenza dal capitale di terzi per finanziare asset a lungo termine, riducendo la flessibilità patrimoniale."
        return "non forniscono elementi sufficienti per un'analisi dettagliata della copertura dell'attivo."

    def get_analisi_gearing(rating):
        if rating == 'A': return "evidenzia un'incidenza minima o nulla del debito finanziario. L'azienda gode di un'elevata autonomia dal sistema bancario, riducendo drasticamente il rischio finanziario e consolidando una posizione di forza e stabilità rispetto alle dinamiche di settore."
        elif rating == 'B': return "evidenzia un livello di indebitamento fisiologico e gestibile. L'azienda utilizza la leva finanziaria in modo equilibrato, mantenendo una struttura patrimoniale complessivamente sostenibile rispetto agli standard del settore."
        elif rating == 'C': return "segnala una forte incidenza del debito rispetto ai mezzi propri. Questo elevato livello di leva finanziaria espone l'azienda a maggiori rischi legati alle variazioni dei tassi di interesse e ne limita l'autonomia dal sistema bancario."
        return "non permette di valutare accuratamente il peso della leva finanziaria."
    
    def get_intro_benchmark_fin(rating):
        if rating == 'A': return "evidenzia una gestione della liquidità ottimale. L'azienda dimostra una notevole capacità di generare risorse e far fronte agli impegni, posizionandosi ai vertici del settore di riferimento:"
        elif rating == 'B': return "evidenzia una gestione della liquidità complessivamente equilibrata. Pur mostrando una struttura finanziaria solida, presenta fisiologici margini di miglioramento rispetto alle medie del settore di riferimento:"
        elif rating == 'C': return "segnala elementi di potenziale criticità nella gestione della liquidità. Le dinamiche finanziarie mostrano margini di miglioramento necessari per garantire una maggiore flessibilità rispetto alle medie del settore di riferimento:"
        return "non consente di esprimere una valutazione completa sulla gestione della liquidità:"

    def get_analisi_rotazione(rating):
        if rating == 'A': return "si distingue per un valore alto, posizionandosi al di sopra degli standard del settore. Questo dato sottolinea un'eccellente efficienza operativa e una notevole capacità di generare flussi di cassa rispetto al capitale investito."
        elif rating == 'B': return "si attesta su valori in linea con le medie del settore. Questo dato evidenzia una buona efficienza operativa e un'adeguata capacità di far ruotare il capitale investito per generare flussi di cassa."
        elif rating == 'C': return "mostra un valore inferiore rispetto agli standard ottimali di settore. Questo dato suggerisce la necessità di migliorare l'efficienza operativa per ottimizzare la generazione di flussi di cassa rispetto al capitale investito."
        return "non fornisce elementi sufficienti per un'analisi dettagliata dell'efficienza operativa."

    def get_analisi_current_ratio(rating):
        if rating == 'A': return "conferma una spiccata solidità finanziaria nel breve periodo. Le attività correnti coprono in modo eccellente le passività correnti, rappresentando una garanzia assoluta per la capacità dell'azienda di far fronte agli impegni finanziari entro l'anno."
        elif rating == 'B': return "indica una sufficiente solidità finanziaria nel breve periodo. Le attività correnti risultano adeguate a coprire le passività correnti, permettendo all'azienda di gestire gli impegni finanziari a breve scadenza."
        elif rating == 'C': return "rappresenta un'area di potenziale attenzione. La copertura delle passività correnti tramite le attività a breve termine risulta limitata, suggerendo cautela nella pianificazione degli impegni finanziari entro l'anno."
        return "non permette di valutare accuratamente la solidità a breve termine."

    def get_analisi_quick_ratio(rating):
        if rating == 'A': return "dimostra un'ottima disponibilità di liquidità immediata. L'azienda è in grado di onorare le scadenze a brevissimo termine senza dover ricorrere allo smobilizzo delle rimanenze, evidenziando una struttura finanziaria estremamente reattiva."
        elif rating == 'B': return "mostra un livello di liquidità immediata accettabile. Pur potendo far fronte agli impegni a breve, l'azienda potrebbe presentare una lieve dipendenza dalla rotazione del magazzino per ottimizzare i flussi di cassa."
        elif rating == 'C': return "costituisce un'area di attenzione. Il valore indica che una quota rilevante della liquidità è vincolata in asset non prontamente convertibili in denaro, segnalando una potenziale fragilità nel rispondere a necessità finanziarie immediate."
        return "non offre dati sufficienti per valutare la reattività di cassa."

    def get_analisi_posizionamento_fin(rating):
        if rating == 'A': return "consolida questo profilo di eccellenza posizionandosi stabilmente nelle fasce più alte delle classifiche di settore, sia su scala nazionale che regionale. L'impresa vanta una struttura finanziaria agile e una notevole flessibilità, che la rendono un punto di riferimento competitivo nel proprio mercato."
        elif rating == 'B': return "mostra un posizionamento intermedio e allineato alle medie di settore all'interno delle classifiche nazionali e regionali. Gli indicatori riflettono una struttura finanziaria resiliente e adatta a mantenere un buon livello di competitività nel contesto locale e generale."
        elif rating == 'C': return "si colloca nelle fasce più basse della distribuzione rispetto al campione settoriale. Questo posizionamento riflette la necessità di strutturare soluzioni per rendere la gestione finanziaria più agile e meno immobilizzata nel capitale circolante, al fine di recuperare competitività locale e nazionale."
        return "presenta un quadro di posizionamento non pienamente valutabile a causa di informazioni incomplete."
    
    def get_analisi_combinata(eco, patr, fin):
        dict_eco = {'A': "un'ottima marginalità", 'B': "una redditività adeguata", 'C': "margini economici da ottimizzare"}
        dict_patr = {'A': "una stabilità patrimoniale d'eccellenza", 'B': "una solida base patrimoniale", 'C': "una struttura patrimoniale da riequilibrare"}
        dict_fin = {'A': "una gestione ottimale della liquidità", 'B': "un equilibrio finanziario stabile", 'C': "una liquidità prontamente disponibile da migliorare"}

        if eco in ['A','B','C'] and patr in ['A','B','C'] and fin in ['A','B','C']:
            return f"riflette un'azienda che poggia su {dict_patr[patr]}, associata a {dict_eco[eco]} e {dict_fin[fin]}."
        return "riflette un'azienda per la quale non è possibile tracciare un profilo combinato completo a causa di dati mancanti."

    def get_descr_fascia_appartenenza(rating):
        if rating == 'A': return "l'eccellenza e il vertice prestazionale"
        elif rating == 'B': return "il cuore solido e ben strutturato"
        elif rating == 'C': return "la fascia che necessita di maggiori interventi di consolidamento"
        return "una porzione non classificabile"

    def get_intro_divario_strutturale(rating_eco):
        if rating_eco == 'A': return "un posizionamento di assoluta leadership gestionale e operativa da parte di"
        elif rating_eco == 'B': return "un generale allineamento fisiologico tra le dinamiche di settore e la gestione di"
        elif rating_eco == 'C': return "un significativo divario strutturale e prestazionale tra l'andamento mediano del comparto e la gestione di"
        return "un quadro prestazionale misto rispetto all'andamento del comparto per"

    def get_analisi_trend_ebitda(rating_eco, mg_ebitda):
        if rating_eco == 'A': return f"ha registrato performance di eccellenza, con un valore al 2024 pari al {mg_ebitda}%. Questo risultato dimostra una straordinaria capacità di assorbire gli shock esterni e mantenere una marginalità operativa nettamente superiore alla mediana settoriale."
        elif rating_eco == 'B': return f"ha mostrato una buona tenuta, con un valore al 2024 pari al {mg_ebitda}%. Questo dato evidenzia una resilienza in linea con la mediana settoriale e una discreta capacità di difendere i margini operativi dalle pressioni macroeconomiche."
        elif rating_eco == 'C': return f"ha registrato evidenti segnali di contrazione o stagnazione, con un valore al 2024 pari al {mg_ebitda}%. Questo risultato risulta modesto e al di sotto della mediana settoriale, segnalando una maggiore difficoltà nell'assorbire i rincari operativi rispetto ai principali player del mercato."
        return "non presenta dati sufficienti per un'analisi storica del posizionamento rispetto al settore."

    def get_asimmetria_ebitda(rating_eco):
        if rating_eco == 'A': return "una spiccata asimmetria positiva nel Margine EBITDA dell'azienda rispetto al benchmark."
        elif rating_eco == 'B': return "un sostanziale allineamento del Margine EBITDA dell'azienda rispetto alle medie del benchmark."
        elif rating_eco == 'C': return "una significativa asimmetria negativa nel Margine EBITDA dell'azienda rispetto al benchmark."
        return "un quadro non pienamente valutabile per il Margine EBITDA."

    def get_confronto_costi_settore(rating_eco):
        if rating_eco == 'A': return "sensibilmente inferiore e più ottimizzata rispetto agli standard mediani del mercato di riferimento"
        elif rating_eco == 'B': return "proporzionata e complessivamente in linea con gli standard mediani del mercato di riferimento"
        elif rating_eco == 'C': return "sensibilmente superiore e più gravosa rispetto agli standard mediani del mercato di riferimento"
        return "non chiaramente determinabile"

    def get_impatto_costi_su_margine(rating_eco):
        if rating_eco == 'A': return "mette in evidenza un'eccellente efficienza nella gestione della spesa corrente, massimizzando il margine lordo e garantendo un netto vantaggio competitivo rispetto ai principali concorrenti"
        elif rating_eco == 'B': return "indica una gestione equilibrata e fisiologica dei costi correnti, che preserva la tenuta del margine lordo pur lasciando ulteriore spazio a future ottimizzazioni"
        elif rating_eco == 'C': return "mette in evidenza una rigidità nella gestione dei costi correnti, la cui incidenza elevata agisce comprimendo il margine lordo e finendo per penalizzare l'azienda rispetto ai principali concorrenti"
        return "non permette di trarre conclusioni definitive sull'efficienza operativa"
    
    def get_analisi_trend_ebit(rating_eco):
        if rating_eco == 'A': return "Questo dato evidenzia una solidità e una progressione estremamente positiva dell'utile operativo netto, a dimostrazione di un rigoroso ed efficiente controllo degli ammortamenti e degli accantonamenti."
        elif rating_eco == 'B': return "Questo dato mostra una sostanziale tenuta dell'utile operativo netto, evidenziando una gestione nel complesso equilibrata e fisiologica degli ammortamenti e degli accantonamenti."
        elif rating_eco == 'C': return "Questo dato evidenzia una dinamica di potenziale contrazione o stagnazione dell'utile operativo netto, riflettendo un'incidenza verosimilmente gravosa degli ammortamenti e degli accantonamenti operativi."
        return "Questo indicatore non presenta dati sufficienti per valutare il trend dell'utile operativo netto."

    def get_confronto_ebit_settore(rating_eco):
        if rating_eco == 'A': return "L'azienda supera ampiamente la redditività operativa del settore, confermando una storica solidità nella gestione dei costi fissi e mantenendo un margine stabilmente superiore rispetto alla mediana dei competitor."
        elif rating_eco == 'B': return "La redditività operativa si mantiene allineata ai valori mediani del settore, confermando una generale adeguatezza nella copertura dei costi operativi fissi rispetto alle dinamiche della concorrenza."
        elif rating_eco == 'C': return "Tale valore resta al di sotto della redditività operativa del comparto, che - al contrario - mantiene mediamente una maggiore solidità strutturale nella copertura dei costi fissi rispetto alle performance dell'azienda analizzata."
        return "Non è tuttavia possibile effettuare un confronto completo con le dinamiche storiche di settore."
    
    def get_sintesi_ebit_mediana(rating_eco):
        if rating_eco == 'A': return "nettamente al di sopra della mediana di settore, segnando un evidente punto di forza per l'impresa"
        elif rating_eco == 'B': return "in sostanziale allineamento con la mediana di settore, confermando un generale equilibrio gestionale"
        elif rating_eco == 'C': return "al di sotto della mediana di settore, rappresentando un'area di potenziale criticità per l'impresa"
        return "in una posizione non chiaramente definibile rispetto alla mediana di settore"

    def get_analisi_verticale_ebit(rating_eco):
        if rating_eco == 'A': return "una virtuosa dinamica dei costi: all'eccellente redditività a livello di EBITDA si affianca un'incidenza degli ammortamenti sul Valore della Produzione perfettamente ottimizzata rispetto agli standard settoriali"
        elif rating_eco == 'B': return "una struttura di costo bilanciata: la redditività a livello di EBITDA è supportata da un'incidenza degli ammortamenti sul Valore della Produzione complessivamente proporzionata agli standard settoriali"
        elif rating_eco == 'C': return "un cosiddetto 'effetto forbice': da una parte si registra una limitata redditività a livello di EBITDA, dall’altra l’incidenza degli ammortamenti sul Valore della Produzione risulta più elevata rispetto agli standard settoriali"
        return "un quadro contabile che necessita di ulteriori approfondimenti sui costi fissi"

    def get_efficienza_struttura_asset(rating_eco):
        if rating_eco == 'A': return "altamente scalabile e particolarmente efficiente nel tradurre gli asset aziendali e il capitale investito in puro valore operativo netto"
        elif rating_eco == 'B': return "adeguatamente dimensionata e sufficientemente efficiente nel tradurre gli asset aziendali in valore operativo netto"
        elif rating_eco == 'C': return "parzialmente sovradimensionata o carente in efficienza nel tradurre gli asset aziendali e il capitale investito in valore operativo netto"
        return "di difficile valutazione in termini di efficienza operativa globale"
    
    def get_analisi_trend_profitto(rating_eco):
        if rating_eco == 'A': return "Questo dato rappresenta l'apice di una gestione d'eccellenza, a conferma della capacità dell'azienda di assorbire efficacemente la pressione fiscale e gli oneri finanziari, consolidando la propria stabilità economica netta nel tempo."
        elif rating_eco == 'B': return "Questo dato riflette una redditività netta nel complesso positiva e stabilizzata, a conferma di un'adeguata gestione della pressione fiscale e degli oneri finanziari nel corso degli ultimi esercizi."
        elif rating_eco == 'C': return "Questo dato riflette una redditività netta ancora contenuta o in potenziale contrazione, segnalando una marcata incidenza degli oneri finanziari, della gestione straordinaria o della pressione fiscale che hanno eroso il risultato di periodo."
        return "Questo dato non risulta sufficiente per tracciare un quadro completo della redditività netta nel tempo."

    def get_confronto_profitto_settore(rating_eco):
        if rating_eco == 'A': return "Il posizionamento risulta di assoluta leadership: l'impresa supera ampiamente le performance mediane dei concorrenti, distinguendosi per una spiccata propensione alla creazione di valore netto."
        elif rating_eco == 'B': return "Il posizionamento risulta pienamente allineato alle performance mediane dei concorrenti, confermando una competitività proporzionata agli standard generali del mercato di riferimento."
        elif rating_eco == 'C': return "Il distacco rispetto ai concorrenti rimane tuttavia un'area di attenzione: l'impresa si colloca al di sotto delle performance mediane del settore, faticando a replicare le traiettorie di consolidamento dei principali player di mercato."
        return "Non è tuttavia possibile stabilire un confronto affidabile con la mediana settoriale."

    def get_prospettiva_redditivita_futura(rating_eco):
        if rating_eco == 'A': return "si affaccia alle sfide future da una posizione di estrema solidità, pronta a cogliere opportunità di espansione."
        elif rating_eco == 'B': return "si affaccia alle sfide future da una posizione di sostanziale equilibrio, con una struttura resiliente ma perfettibile."
        elif rating_eco == 'C': return "si ritrova a fronteggiare una fase di potenziale fragilità, richiedendo interventi tempestivi per invertire il trend reddituale."
        return "affronta il prossimo esercizio con un quadro reddituale in via di definizione."

    def get_posizionamento_margine_profitto_fine(rating_eco):
        if rating_eco == 'A': return "ai vertici del mercato, riflettendo una spiccata propensione alla creazione di valore netto"
        elif rating_eco == 'B': return "su livelli intermedi, in generale allineamento con le performance dei principali competitor"
        elif rating_eco == 'C': return "su livelli di attenzione rispetto ai competitor, limitando di fatto i margini di manovra dell'impresa"
        return "su livelli non pienamente quantificabili"

    def get_sintesi_bilancio_finale(rating_eco):
        if rating_eco == 'A': return "i frutti di una gestione oculata e altamente strategica"
        elif rating_eco == 'B': return "i risultati di una gestione nel complesso ordinaria e bilanciata"
        elif rating_eco == 'C': return "le conseguenze di inefficienze gestionali o rigidità strutturali non ancora risolte"
        return "un quadro contabile misto"

    def get_impatto_gestione_caratteristica(rating_eco):
        if rating_eco == 'A': return "ottimale, massimizzando l'efficienza dei costi operativi e degli ammortamenti"
        elif rating_eco == 'B': return "proporzionata, evidenziando una fisiologica incidenza dei costi operativi e degli ammortamenti"
        elif rating_eco == 'C': return "eccessivamente gravosa, penalizzando il risultato a causa di crescenti costi operativi o ammortamenti pesanti"
        return "variabile"

    def get_impatto_oneri_accessori(rating_eco):
        if rating_eco == 'A': return "risulta ampiamente sostenibile, permettendo all'azienda di produrre un utile di assoluto rilievo"
        elif rating_eco == 'B': return "risulta gestibile, consentendo all'azienda di produrre un margine netto soddisfacente"
        elif rating_eco == 'C': return "aggiunge un ulteriore elemento di pressione, che rischia di compromettere la capacità dell'azienda di produrre un margine netto competitivo"
        return "non permette di esprimere un giudizio definitivo"
    
    def get_interpretazione_risultato_eco(rating_eco):
        if rating_eco == 'A': return "riflette una configurazione ottimale e strutturata, confermando la piena solidità del modello di business."
        elif rating_eco == 'B': return "indica un profilo complessivamente stabile, pur evidenziando fisiologiche aree suscettibili di ulteriore ottimizzazione."
        elif rating_eco == 'C': return "non può essere attribuito a fattori puramente transitori, ma evidenzia uno squilibrio di natura strutturale."
        return "restituisce un quadro di non immediata e univoca interpretazione."

    def get_stato_struttura_operativa(rating_eco):
        if rating_eco == 'A': return "altamente performante, caratterizzata da spiccata flessibilità ed efficienza"
        elif rating_eco == 'B': return "adeguata e resiliente, caratterizzata da un generale equilibrio gestionale"
        elif rating_eco == 'C': return "inadeguata rispetto agli standard competitivi, caratterizzata da rigidità ed inefficienza"
        return "in via di definizione rispetto ai parametri medi"

    def get_gestione_costi_ricavi_core(rating_eco):
        if rating_eco == 'A': return "gestisce in modo eccellente i costi operativi esterni, potendo contare su ricavi core ampiamente in grado di coprire i costi fissi, gli ammortamenti e gli oneri finanziari, generando un surplus significativo"
        elif rating_eco == 'B': return "mantiene un'incidenza proporzionata dei costi operativi, con ricavi core in grado di coprire in modo sufficiente i costi fissi, gli ammortamenti e gli oneri finanziari, garantendo la continuità aziendale"
        elif rating_eco == 'C': return "sostiene costi operativi esterni significativamente elevati e deve fronteggiare un onere sproporzionato di costi fissi, come ammortamenti e oneri finanziari, che i ricavi core non sono più in grado di coprire adeguatamente"
        return "presenta un assorbimento dei costi da parte dei ricavi parzialmente tracciabile"

    def get_implicazione_finale_eco(rating_eco):
        if rating_eco == 'A': return "Il posizionamento ai vertici del mercato rappresenta una chiara conferma della validità delle strategie intraprese: mantenendo questo livello di ottimizzazione dei costi e redditività degli asset, l'azienda si garantisce una leadership economica duratura e sostenibile rispetto ai concorrenti."
        elif rating_eco == 'B': return "Il posizionamento nella fascia media del mercato rappresenta una base sicura: attraverso mirati interventi strategici volti a ottimizzare ulteriormente i costi operativi e migliorare la redditività degli asset, l'azienda potrà puntare a un consolidamento superiore nel proprio mercato di riferimento."
        elif rating_eco == 'C': return "Il posizionamento in questa fascia rappresenta un chiaro campanello d'allarme: senza un intervento strategico urgente volto a tagliare i costi operativi e migliorare la redditività degli asset, l'azienda rischia di restare intrappolata in una condizione di insostenibilità economica rispetto ai livelli minimi richiesti dal mercato."
        return "L'assenza di dati completi suggerisce cautela nell'esprimere un verdetto definitivo sulla continuità economica di lungo periodo."

    def get_sintesi_quadriennio_patr(rating_patr):
        if rating_patr == 'A': return "un profilo di notevole solidità e indipendenza finanziaria, ponendo l'impresa nettamente al di sopra dei benchmark settoriali"
        elif rating_patr == 'B': return "un profilo di sostanziale equilibrio e adeguatezza patrimoniale, in linea con le dinamiche mediane dei benchmark settoriali"
        elif rating_patr == 'C': return "un profilo di vulnerabilità e forte dipendenza finanziaria, collocando l'impresa al di sotto dei principali benchmark settoriali"
        return "un quadro di equilibrio patrimoniale in via di definizione"

    def get_analisi_copertura_attivo_fisso(rating_patr):
        if rating_patr == 'A': return "l'impresa finanzia interamente, o quasi, le proprie immobilizzazioni e i propri investimenti strutturali tramite il Capitale Proprio, svincolandosi dalle logiche del debito"
        elif rating_patr == 'B': return "l'impresa finanzia una quota adeguata delle proprie immobilizzazioni tramite il Capitale Proprio, bilanciando in modo fisiologico il ricorso al capitale di terzi"
        elif rating_patr == 'C': return "l'impresa presenta un deficit di copertura strutturale, finanziando gran parte delle proprie immobilizzazioni tramite il capitale di terzi ed esponendosi maggiormente ai rischi"
        return "la struttura di copertura del capitale non è pienamente valutabile"

    def get_confronto_struttura_settore(rating_patr):
        if rating_patr == 'A': return "il posizionamento dell'azienda si conferma di assoluta eccellenza. Mentre un'ampia fetta del mercato continua a finanziare l'acquisto di asset strutturali attraverso debito a lungo termine, l'azienda utilizza il Patrimonio Netto come scudo e leva primaria di crescita"
        elif rating_patr == 'B': return "il posizionamento dell'azienda risulta allineato ai valori mediani. In un mercato che sta gradualmente ottimizzando il rapporto tra debito a lungo termine e mezzi propri, l'impresa partecipa attivamente a questo processo di consolidamento"
        elif rating_patr == 'C': return "il divario rispetto alla mediana di mercato appare evidente. Mentre il settore avvia processi di riduzione del debito e rafforzamento patrimoniale, l'azienda sconta ancora una marcata dipendenza dalla leva finanziaria a lungo termine"
        return "non è tuttavia possibile tracciare un confronto storico completo con il settore"

    def get_analisi_soglia_struttura1(rating_patr):
        if rating_patr == 'A': return "Tale valore evidenzia una performance di capitalizzazione di assoluto rilievo, superando ampiamente le soglie ideali di sicurezza strutturale (>1)."
        elif rating_patr == 'B': return "Tale valore evidenzia una struttura di capitale adeguata, posizionandosi in prossimità delle soglie di sicurezza e garantendo una fisiologica solidità di base."
        elif rating_patr == 'C': return "Tale valore si colloca al di sotto delle soglie ottimali di sicurezza, segnalando una potenziale tensione nella struttura di capitalizzazione dell'impresa."
        return "Tale indicatore non permette di definire univocamente il posizionamento rispetto alle soglie di sicurezza."

    def get_confronto_mediana_struttura1(rating_patr):
        if rating_patr == 'A': return "Il posizionamento risulta nettamente superiore rispetto alla media del mercato di riferimento."
        elif rating_patr == 'B': return "Il posizionamento risulta sostanzialmente allineato alle dinamiche mediane del mercato di riferimento."
        elif rating_patr == 'C': return "Il posizionamento risulta inferiore e maggiormente esposto rispetto alle medie del mercato di riferimento."
        return "Il confronto strutturale con la mediana settoriale non risulta pienamente valutabile."

    def get_implicazione_copertura_attivo(rating_patr):
        if rating_patr == 'A': return "non solo è in grado di coprire compiutamente il valore del proprio attivo fisso con il capitale proprio, ma si pone in una posizione di elevata immunità rispetto alle restrizioni creditizie e alle preoccupazioni di insolvenza che affliggono i concorrenti."
        elif rating_patr == 'B': return "è in grado di garantire un'adeguata copertura degli investimenti a lungo termine, mantenendo un'esposizione al credito gestibile e preservando una serena continuità operativa."
        elif rating_patr == 'C': return "finanzia una quota rilevante del proprio attivo fisso ricorrendo a capitale di terzi, esponendosi maggiormente alle fluttuazioni dei tassi e alle restrizioni creditizie rispetto ai concorrenti dotati di una base patrimoniale più solida."
        return "presenta un quadro di copertura degli investimenti in via di consolidamento."
    
    def get_analisi_trend_struttura2(rating_patr):
        if rating_patr == 'A': return "Questo dato conferma un'eccellente capacità di copertura dell'attivo fisso tramite il capitale proprio e le passività consolidate, delineando una struttura finanziaria altamente solida."
        elif rating_patr == 'B': return "Questo dato indica un'adeguata capacità di copertura dell'attivo fisso tramite le fonti permanenti di capitale, delineando una struttura finanziaria complessivamente bilanciata."
        elif rating_patr == 'C': return "Questo dato evidenzia una parziale vulnerabilità nella copertura dell'attivo fisso, segnalando un ricorso potenzialmente eccessivo a fonti di finanziamento a breve termine per sostenere investimenti duraturi."
        return "Questo dato non risulta sufficiente per elaborare un'analisi completa sulla copertura a lungo termine."

    def get_confronto_settore_struttura2(rating_patr):
        if rating_patr == 'A': return "Il valore risulta significativamente più elevato rispetto alla mediana settoriale, confermando un netto vantaggio competitivo nella gestione delle fonti di finanziamento."
        elif rating_patr == 'B': return "Il valore si mantiene in sostanziale allineamento con la mediana settoriale, rispecchiando le dinamiche fisiologiche del mercato di riferimento."
        elif rating_patr == 'C': return "Il valore risulta inferiore rispetto alla mediana settoriale, evidenziando una maggiore esposizione dell'impresa rispetto ai principali competitor."
        return "Un confronto diretto con la mediana di settore non è al momento elaborabile."

    def get_conclusione_struttura2(rating_patr):
        if rating_patr == 'A': return "riflette una prudente e virtuosa strategia di capitalizzazione che protegge l'azienda da potenziali tensioni di liquidità a lungo termine e dai rischi di rifinanziamento"
        elif rating_patr == 'B': return "riflette una strategia di capitalizzazione adeguata, che garantisce la continuità operativa limitando l'esposizione ai rischi di rifinanziamento nel lungo periodo"
        elif rating_patr == 'C': return "suggerisce la necessità di un riequilibrio delle fonti di finanziamento, al fine di mitigare l'esposizione alle restrizioni del credito e proteggere l'azienda dai rischi di rifinanziamento"
        return "rimane in attesa di un consolidamento dei dati storici"

    def get_evoluzione_vantaggio_competitivo(rating_patr):
        if rating_patr == 'A': return "riesce a consolidare un vantaggio competitivo strutturale di altissimo profilo rispetto ai principali operatori del comparto."
        elif rating_patr == 'B': return "mantiene una solida e fisiologica posizione competitiva, ben allineata alle dinamiche di stabilizzazione del comparto."
        elif rating_patr == 'C': return "affronta una fase di transizione sfidante, evidenziando la necessità di colmare un divario strutturale rispetto ai principali operatori del comparto."
        return "mostra un posizionamento competitivo in fase di definizione."

    def get_resilienza_credit_crunch(rating_patr):
        if rating_patr == 'A': return "appare altamente resiliente alle fluttuazioni e agli effetti delle restrizioni creditizie, grazie a un saldo ampiamente positivo di fonti stabili"
        elif rating_patr == 'B': return "mostra una buona tenuta rispetto agli effetti delle dinamiche creditizie, supportata da un adeguato e proporzionato equilibrio delle fonti stabili"
        elif rating_patr == 'C': return "risulta maggiormente esposto agli effetti delle restrizioni creditizie, evidenziando la chiara necessità di incrementare la quota di fonti stabili"
        return "presenta caratteristiche di stabilità in via di assestamento"

    def get_flessibilita_strategica_asset(rating_patr):
        if rating_patr == 'A': return "un'elevata flessibilità strategica nella gestione, nel rinnovo e nello sviluppo dei propri asset operativi"
        elif rating_patr == 'B': return "una flessibilità strategica sufficiente e idonea per la gestione ordinaria e il rinnovamento dei propri asset operativi"
        elif rating_patr == 'C': return "limitati margini di flessibilità strategica per la gestione e lo sviluppo di lungo termine dei propri asset operativi"
        return "un quadro di flessibilità gestionale da monitorare nel tempo"
    
    def get_intro_analisi_gearing(rating_patr):
        if rating_patr == 'A': return "emerge un quadro di eccellente indipendenza finanziaria."
        elif rating_patr == 'B': return "emerge un quadro di fisiologico ricorso al capitale di terzi."
        elif rating_patr == 'C': return "emerge una marcata esposizione strutturale verso il capitale di terzi."
        return "emerge un quadro di indipendenza finanziaria in via di accertamento."

    def get_andamento_storico_gearing(rating_patr):
        if rating_patr == 'A': return "Questo dato riflette un livello di indebitamento oneroso nullo o estremamente contenuto, a testimonianza di una solida politica di autofinanziamento."
        elif rating_patr == 'B': return "Questo dato riflette un livello di indebitamento sostenibile e bilanciato, in grado di supportare l'operatività senza appesantire in modo eccessivo la struttura."
        elif rating_patr == 'C': return "Questo dato riflette un livello di indebitamento elevato, segnalando una forte dipendenza dalla leva finanziaria per il sostegno degli investimenti o dell'operatività."
        return "Tale indicatore necessita di ulteriori approfondimenti per valutarne la portata storica."

    def get_confronto_gearing_settore(rating_patr):
        if rating_patr == 'A': return "l'esposizione debitoria dell'impresa risulta nettamente inferiore e strutturalmente più sicura rispetto alle medie del comparto."
        elif rating_patr == 'B': return "l'esposizione debitoria dell'impresa risulta in sostanziale allineamento con la fisiologica leva finanziaria utilizzata dalle aziende del comparto."
        elif rating_patr == 'C': return "l'esposizione debitoria dell'impresa risulta sensibilmente superiore rispetto alle medie di settore, delineando un profilo di rischio più elevato."
        return "non è possibile tracciare un parallelismo completo con i livelli di indebitamento del settore."

    def get_reazione_contesto_gearing(rating_patr):
        if rating_patr == 'A': return "rappresenta un'eccellenza, distinguendosi per un'assenza quasi totale di debito finanziario oneroso e per un'incredibile immunità dalle fluttuazioni dei tassi d'interesse"
        elif rating_patr == 'B': return "mantiene una gestione ragionata, assorbendo le variazioni dei tassi d'interesse all'interno di un equilibrio patrimoniale controllato e mitigando le tensioni sui rimborsi"
        elif rating_patr == 'C': return "risulta particolarmente vulnerabile, subendo in modo diretto l'impatto dei tassi variabili e rendendo prioritaria una rigorosa strategia di rientro e consolidamento del debito"
        return "mostra dinamiche di reazione al costo del denaro in via di consolidamento"
    
    def get_sintesi_valore_gearing(rating_patr):
        if rating_patr == 'A': return "svela la chiave definitiva dell'eccellenza strutturale aziendale."
        elif rating_patr == 'B': return "conferma il generale stato di salute e l'equilibrio strutturale dell'impresa."
        elif rating_patr == 'C': return "rappresenta il nodo critico dell'attuale struttura patrimoniale dell'impresa."
        return "non permette di trarre una conclusione definitiva sulla solidità aziendale."

    def get_analisi_rischio_default(rating_patr):
        if rating_patr == 'A': return "un profilo di rischio finanziario praticamente inesistente, blindando l'azienda da potenziali shock esterni o restrizioni creditizie."
        elif rating_patr == 'B': return "un profilo di rischio finanziario moderato e fisiologico, gestibile attraverso i normali flussi operativi senza compromettere la continuità aziendale."
        elif rating_patr == 'C': return "un profilo di rischio finanziario elevato, rendendo l'impresa vulnerabile a potenziali shock esterni, rincari dei tassi o richieste di rientro da parte delle banche."
        return "un profilo di rischio finanziario in via di accertamento."

    def get_conclusione_autonomia_finanziaria(rating_patr):
        if rating_patr == 'A': return "si distingue per l'assenza di un indebitamento strutturale significativo, garantendosi una piena e duratura autonomia finanziaria."
        elif rating_patr == 'B': return "raggiunge un compromesso sostenibile tra mezzi propri e capitale di terzi, preservando una sufficiente autonomia finanziaria."
        elif rating_patr == 'C': return "evidenzia un'esposizione marcata verso il sistema creditizio, limitando fortemente la propria autonomia finanziaria e le prospettive di investimento indipendente."
        return "mostra un quadro di autonomia finanziaria ancora da consolidare."
    
    def get_sintesi_finale_patr(rating_patr):
        if rating_patr == 'A': return "rappresenta un esempio di eccellenza e riflette una gestione aziendale prudente, solida e ampiamente capitalizzata, capace di sostenere agevolmente l'operatività corrente."
        elif rating_patr == 'B': return "riflette una gestione aziendale nel complesso equilibrata e adeguatamente capitalizzata, con una solida tenuta strutturale di base."
        elif rating_patr == 'C': return "evidenzia alcune criticità strutturali, riflettendo una capitalizzazione non ottimale che necessita di interventi di riequilibrio per garantire maggiore sicurezza."
        return "non permette di formulare un giudizio definitivo a causa di dati incompleti."

    def get_impatto_rischio_sistemico(rating_patr):
        if rating_patr == 'A': return "mitiga quasi totalmente il rischio sistemico, garantendo un’elevata protezione e una corazza inespugnabile contro eventuali instabilità nel mercato del credito"
        elif rating_patr == 'B': return "mantiene il rischio sistemico a livelli fisiologici e gestibili, offrendo una protezione adeguata contro le normali fluttuazioni del mercato creditizio"
        elif rating_patr == 'C': return "espone l'impresa a un maggiore rischio sistemico, riducendo significativamente le difese interne contro eventuali restrizioni creditizie e shock di mercato"
        return "presenta un livello di esposizione al rischio non chiaramente calcolabile"

    def get_posizionamento_competitivo_patr(rating_patr):
        if rating_patr == 'A': return "si distingue come un modello virtuoso e un punto di riferimento di solidità per l’intero comparto,"
        elif rating_patr == 'B': return "si inserisce stabilmente nella media del settore, dimostrando buona resilienza,"
        elif rating_patr == 'C': return "affronta una sfida cruciale per riallinearsi agli standard di sicurezza del mercato,"
        return "presenta un quadro di posizionamento da consolidare,"
    
    def get_sintesi_modello_operativo_fin(rating_fin):
        if rating_fin == 'A': return "un modello operativo di assoluta eccellenza, caratterizzato da una straordinaria velocità di rotazione del capitale e da una gestione della liquidità ampiamente superiore alla mediana del settore."
        elif rating_fin == 'B': return "un modello operativo nel complesso equilibrato e in linea con il settore, mostrando una fisiologica e adeguata capacità di far fronte agli impegni a breve termine."
        elif rating_fin == 'C': return "un modello operativo che si discosta dalle efficienze ottimali di settore, evidenziando alcune rigidità nella rotazione del capitale investito e flussi operativi meno dinamici."
        return "un modello operativo la cui efficienza finanziaria è ancora in via di definizione."

    def get_analisi_trend_current_ratio(rating_fin):
        if rating_fin == 'A': return "riflette un'eccellente capacità di copertura delle passività correnti attraverso l'attivo circolante, posizionandosi stabilmente al di sopra dei livelli di sicurezza e della mediana di settore."
        elif rating_fin == 'B': return "riflette una buona capacità di copertura delle passività correnti attraverso l'attivo circolante, garantendo la solvibilità aziendale in sostanziale allineamento con gli standard del mercato."
        elif rating_fin == 'C': return "denota una potenziale vulnerabilità nella copertura delle passività correnti attraverso l'attivo circolante, posizionandosi su livelli di attenzione rispetto agli standard di sicurezza richiesti."
        return "non risulta sufficiente per elaborare una valutazione storica della solvibilità a breve termine."

    def get_reazione_contesto_liquidita(rating_fin):
        if rating_fin == 'A': return "ha dimostrato una straordinaria resilienza, difendendo i propri margini di liquidità e confermando una struttura finanziaria pressoché immune alle dinamiche di erosione che hanno colpito il comparto"
        elif rating_fin == 'B': return "ha saputo gestire l'impatto di tali dinamiche, assorbendo le pressioni inflattive con una fisiologica flessione della liquidità che non ne ha comunque compromesso la continuità operativa"
        elif rating_fin == 'C': return "si è trovata maggiormente esposta a questa fase di erosione, registrando un assorbimento della liquidità più marcato rispetto alla media, il che rende prioritario un tempestivo ripristino delle scorte di cassa"
        return "ha mostrato dinamiche di assorbimento della cassa in fase di assestamento"
    
    def get_confronto_current_ratio_mediana(rating_fin):
        if rating_fin == 'A': return "appare di assoluta eccellenza, con un Current Ratio che si posiziona stabilmente al di sopra del benchmark di riferimento"
        elif rating_fin == 'B': return "appare complessivamente equilibrata, con un Current Ratio che si mantiene allineato al benchmark di riferimento"
        elif rating_fin == 'C': return "appare più critica, con un Current Ratio che si colloca al di sotto delle soglie ottimali, allontanandosi dal benchmark di riferimento"
        return "restituisce un quadro di posizionamento non pienamente definito"

    def get_rapporto_attivita_passivita_brevi(rating_fin):
        if rating_fin == 'A': return "le attività correnti dell'azienda coprano in modo eccellente le passività a breve termine, garantendo un'ampia flessibilità finanziaria e azzerando le tensioni di cassa"
        elif rating_fin == 'B': return "le attività correnti dell'azienda siano sufficienti a coprire le passività a breve termine, mantenendo una fisiologica condizione di stabilità finanziaria"
        elif rating_fin == 'C': return "le passività a breve termine dell'azienda premano in modo significativo sulle sue attività correnti, portandola in una condizione di potenziale tensione finanziaria"
        return "la proporzione tra attività e passività a breve termine risulti in fase di consolidamento"

    def get_capacita_generazione_liquidita(rating_fin):
        if rating_fin == 'A': return "dimostra una straordinaria capacità di generare e mantenere una liquidità abbondante, affrontando in totale serenità gli obblighi finanziari a breve termine"
        elif rating_fin == 'B': return "dimostra un'adeguata capacità di mantenere la liquidità necessaria per far fronte agli obblighi finanziari a breve termine senza particolari criticità operative"
        elif rating_fin == 'C': return "riscontra oggettive difficoltà nel generare o mantenere un buffer di liquidità sufficiente per far fronte in totale sicurezza agli obblighi finanziari a breve termine"
        return "mostra un andamento dei flussi di liquidità da monitorare"
    
    def get_analisi_quick_ratio_soglia(rating_fin):
        if rating_fin == 'A': return "si posiziona nettamente al di sopra della soglia ideale di sicurezza, evidenziando un'eccellente reattività di cassa e superando ampiamente le performance della mediana settoriale"
        elif rating_fin == 'B': return "si mantiene in prossimità delle soglie di sicurezza, allineandosi in modo fisiologico e bilanciato alle dinamiche della mediana settoriale"
        elif rating_fin == 'C': return "si colloca al di sotto della soglia ottimale di sicurezza e della mediana settoriale, mettendo in luce un'area di potenziale criticità e attenzione per la società"
        return "restituisce un quadro di posizionamento non pienamente quantificabile"

    def get_implicazione_liquidita_immediata(rating_fin):
        if rating_fin == 'A': return "la disponibilità di liquidità immediata risulta eccellente e del tutto indipendente dallo smobilizzo del magazzino, garantendo all'impresa una struttura finanziaria estremamente elastica e superiore a quella dei concorrenti"
        elif rating_fin == 'B': return "la disponibilità di liquidità immediata risulta adeguata al ciclo operativo aziendale, mostrando una gestione fisiologica e controllabile delle rimanenze per far fronte agli impegni a brevissimo termine"
        elif rating_fin == 'C': return "una quota rilevante dell'attivo circolante risulta vincolata in rimanenze; pertanto, la disponibilità di liquidità immediata dipenderà strettamente dal ciclo delle vendite e dallo smobilizzo del magazzino, denotando una struttura finanziaria meno elastica rispetto ai concorrenti"
        return "la reale disponibilità di liquidità a brevissimo termine rimane da monitorare"
    
    def get_andamento_liquidita_primaria(rating_fin):
        if rating_fin == 'A': return "la straordinaria capacità dell'azienda di consolidare e incrementare la propria liquidità primaria"
        elif rating_fin == 'B': return "la fisiologica capacità dell'azienda di mantenere stabile la propria liquidità primaria"
        elif rating_fin == 'C': return "la difficoltà dell'azienda nel ricostituire o mantenere la propria liquidità primaria ai livelli ottimali"
        return "le dinamiche di assestamento della liquidità primaria"

    def get_confronto_quick_ratio_settore(rating_fin):
        if rating_fin == 'A': return "ben al di sopra della mediana di settore"
        elif rating_fin == 'B': return "in sostanziale allineamento con la mediana di settore"
        elif rating_fin == 'C': return "al di sotto della mediana di settore"
        return "su livelli non pienamente confrontabili con la mediana di settore"

    def get_copertura_debiti_breve_quick(rating_fin):
        if rating_fin == 'A': return "sono più che sufficienti a coprire agevolmente i debiti a breve scadenza, garantendo un ampio e solido margine di manovra"
        elif rating_fin == 'B': return "risultano adeguati a coprire in modo sufficiente i debiti a breve scadenza, preservando il necessario equilibrio di cassa"
        elif rating_fin == 'C': return "non risultano pienamente sufficienti a coprire i debiti a breve scadenza in modo autonomo, creando dipendenza dal magazzino"
        return "presentano un quadro di copertura in via di definizione"

    def get_interpretazione_flussi_cassa_quick(rating_fin):
        if rating_fin == 'A': return "rappresenta un chiaro segnale di eccellente efficienza nella gestione degli incassi e dei flussi operativi, ponendo l'azienda al riparo da tensioni di liquidità nel breve periodo"
        elif rating_fin == 'B': return "conferma una gestione ordinata degli incassi e dei flussi operativi, limitando di fatto l'esposizione ai rischi di illiquidità nel breve periodo"
        elif rating_fin == 'C': return "rappresenta un potenziale segnale di inefficienza nella gestione degli incassi o di un anomalo assorbimento di cassa per l'operatività corrente, esponendo l'azienda a un tangibile rischio di illiquidità nel breve periodo"
        return "necessita di ulteriori approfondimenti sulle dinamiche dei flussi di cassa"
    
    def get_confronto_rotazione_mediana(rating_fin):
        if rating_fin == 'A': return "si colloca su livelli di assoluta eccellenza, risultando nettamente superiore alla mediana settoriale e confermando una velocità di rotazione invidiabile."
        elif rating_fin == 'B': return "si colloca in sostanziale allineamento con la mediana settoriale, riflettendo una velocità di rotazione adeguata e fisiologica per il comparto."
        elif rating_fin == 'C': return "si colloca al di sotto della mediana settoriale, evidenziando una velocità di rotazione del capitale più lenta rispetto ai principali competitor."
        return "non risulta facilmente confrontabile con le medie di mercato."

    def get_interpretazione_modello_rotazione(rating_fin):
        if rating_fin == 'A': return "denota un modello di business estremamente agile ed efficiente, capace di generare volumi di fatturato molto elevati rispetto alla dotazione patrimoniale investita."
        elif rating_fin == 'B': return "denota un modello di business equilibrato, che riesce a bilanciare la gestione degli asset con la generazione di ricavi in modo coerente con le aspettative del mercato."
        elif rating_fin == 'C': return "suggerisce un modello di business caratterizzato da una gestione degli asset potenzialmente rigida, dove la generazione di ricavi fatica a tenere il passo con l'ammontare del capitale investito."
        return "richiede un ulteriore consolidamento storico per valutare il modello di business."
    
    def get_analisi_anomalia_rotazione(rating_fin):
        if rating_fin == 'A': return "tale valore evidenzia una spiccata efficienza commerciale e una configurazione operativa altamente ottimizzata."
        elif rating_fin == 'B': return "tale valore riflette un'efficienza commerciale adeguata e una configurazione aziendale allineata agli standard fisiologici."
        elif rating_fin == 'C': return "tale valore evidenzia una rigidità commerciale e una configurazione operativa meno dinamica rispetto ai competitor."
        return "tale valore restituisce un quadro di efficienza in via di definizione."

    def get_rapporto_capitale_fatturato(rating_fin):
        if rating_fin == 'A': return "risulti particolarmente contenuto e ottimizzato in relazione all'elevato volume d'affari generato"
        elif rating_fin == 'B': return "risulti proporzionato ed equilibrato rispetto al volume d'affari sviluppato"
        elif rating_fin == 'C': return "risulti eccessivamente gravoso e sproporzionato rispetto al limitato volume d'affari generato"
        return "non sia ancora pienamente stabilizzato rispetto ai volumi di vendita"

    def get_motivazione_strutturale_rotazione(rating_fin):
        if rating_fin == 'A': return "un modello di business agile, capace di massimizzare la resa degli asset operativi storici o di sfruttare cicli di vendita estremamente veloci."
        elif rating_fin == 'B': return "un modello di business consolidato, in cui la dotazione patrimoniale supporta coerentemente i cicli di vendita senza generare particolari inefficienze."
        elif rating_fin == 'C': return "un modello di business appesantito da asset scarsamente produttivi o da un ciclo delle vendite rallentato, che necessita di mirati interventi di ottimizzazione."
        return "un modello di business in fase di transizione e assestamento."
    
    def get_sintesi_finale_fin(rating_fin):
        if rating_fin == 'A': return "uno stato di salute eccellente, caratterizzato da un'abbondante liquidità e un'ottimale rotazione del capitale investito"
        elif rating_fin == 'B': return "un quadro complessivamente equilibrato, caratterizzato da una liquidità adeguata e una fisiologica rotazione del capitale"
        elif rating_fin == 'C': return "un profondo squilibrio strutturale, caratterizzato da un evidente deterioramento degli indicatori di liquidità nonostante i volumi generati"
        return "un quadro di liquidità e rotazione in fase di accertamento"

    def get_motivazione_rating_fin(rating_fin):
        if rating_fin == 'A': return "la straordinaria capacità dell'azienda di mantenere livelli ottimali di Current Ratio e Quick Ratio, garantendo una totale copertura dei debiti a breve termine e azzerando i rischi di illiquidità"
        elif rating_fin == 'B': return "la buona capacità dell'azienda di mantenere livelli sufficienti di Current Ratio e Quick Ratio, assicurando la copertura dei debiti a breve e mitigando i rischi operativi"
        elif rating_fin == 'C': return "l'incapacità dell'azienda di mantenere livelli adeguati di Current Ratio e Quick Ratio per coprire i debiti a breve termine, evidenziando una rigidità finanziaria che i modelli di valutazione penalizzano severamente"
        return "le fisiologiche dinamiche di incasso e pagamento aziendali"

    def get_gestione_tesoreria_fin(rating_fin):
        if rating_fin == 'A': return "si distingue per una gestione virtuosa della tesoreria: l'operatività genera flussi di cassa abbondanti e costanti, fornendo ampie risorse per far fronte alle obbligazioni correnti senza alcuno stress"
        elif rating_fin == 'B': return "mantiene un controllo adeguato sulla tesoreria: la gestione operativa genera flussi sufficienti a sostenere gli impegni a breve, mantenendo un corretto allineamento temporale tra incassi e uscite"
        elif rating_fin == 'C': return "si trova ad affrontare una forte pressione sulla tesoreria: la gestione operativa assorbe liquidità invece di generarla, prosciugando le risorse necessarie per far fronte alle obbligazioni correnti e creando un pericoloso disallineamento temporale"
        return "mostra un andamento dei flussi di cassa da stabilizzare"

    def get_priorita_strategica_fin(rating_fin):
        if rating_fin == 'A': return "non tanto il risanamento, ormai ampiamente superato, quanto il mantenimento di queste eccellenti performance di cassa e la massimizzazione del rendimento per gli investimenti futuri"
        elif rating_fin == 'B': return "il progressivo rafforzamento del capitale circolante e la costante ottimizzazione dei cicli di incasso, al fine di blindare ulteriormente l'indipendenza finanziaria a breve termine"
        elif rating_fin == 'C': return "non tanto la mera efficienza commerciale, quanto il recupero urgente del capitale circolante, la dilazione del debito a breve e il rafforzamento vitale delle riserve liquide"
        return "il monitoraggio costante dei flussi di cassa nel breve termine"
    
    def get_sintesi_profilo_integrato(rating_comb):
        if rating_comb == 'A': return "altamente competitivo e strutturalmente solido."
        elif rating_comb == 'B': return "complessivamente equilibrato e resiliente."
        else: return "caratterizzato da vulnerabilità strutturali e operative."

    def get_sintesi_posizionamento_lungo_periodo(rating_comb):
        if rating_comb == 'A': return "leader, capace di affrontare con sicurezza le sfide di lungo periodo."
        elif rating_comb == 'B': return "stabile, in grado di mantenere la propria posizione competitiva nel tempo."
        else: return "esposta, che necessita di interventi per garantire la continuità nel lungo periodo."

    def get_conclusione_patrimoniale(rating_patr):
        if rating_patr == 'A': return "l'azienda dimostra un'eccellente capacità di coprire gli investimenti con mezzi propri."
        elif rating_patr == 'B': return "l'azienda mantiene un adeguato bilanciamento tra mezzi propri e capitale di terzi."
        else: return "emerge la necessità di ridurre l'esposizione debitoria per riequilibrare la struttura."

    def get_conclusione_economica(rating_eco):
        if rating_eco == 'A': return "confermano una straordinaria capacità di generare valore e ottimizzare i costi."
        elif rating_eco == 'B': return "mostrano una redditività soddisfacente e in linea con le aspettative del mercato."
        else: return "evidenziano l'urgenza di un contenimento dei costi per recuperare marginalità."

    def get_conclusione_finanziaria_dettaglio(rating_fin):
        if rating_fin == 'A': return "garantisce un'ottima elasticità di cassa e sicurezza nei pagamenti a breve."
        elif rating_fin == 'B': return "assicura una solvibilità adeguata per la gestione operativa corrente."
        else: return "segnala possibili tensioni di liquidità da monitorare con attenzione."

    def get_raccomandazione_finale(rating_comb):
        if rating_comb == 'A': return "possiede tutte le carte in regola per puntare a un'ulteriore espansione nel proprio settore."
        elif rating_comb == 'B': return "dovrebbe concentrarsi sul consolidamento dei propri punti di forza e sull'efficientamento delle aree meno performanti."
        else: return "dovrebbe adottare tempestivamente misure di risanamento strutturale e finanziario per invertire il trend."




    context['descr_rating_tot'] = get_testo_totale(context['rating_comb'])
    context['descr_rating_eco'] = get_testo_eco(context['rating_eco'])
    context['descr_rating_patr'] = get_testo_patr(context['rating_patr'])
    context['descr_rating_fin'] = get_testo_fin(context['rating_fin'])
    context['descr_sintesi'] = get_testo_sintesi(context['rating_comb'])
    context['intro_benchmark_eco'] = get_intro_benchmark_eco(context['rating_eco'])
    context['analisi_margini_operativi'] = get_analisi_margini_operativi(context['rating_eco'])
    context['analisi_margine_profitto'] = get_analisi_margine_profitto(context['rating_eco'])
    context['intro_benchmark_patr'] = get_intro_benchmark_patr(context['rating_patr'])
    context['analisi_indici_struttura'] = get_analisi_indici_struttura(context['rating_patr'])
    context['analisi_gearing'] = get_analisi_gearing(context['rating_patr'])
    context['intro_benchmark_fin'] = get_intro_benchmark_fin(context['rating_fin'])
    context['analisi_rotazione'] = get_analisi_rotazione(context['rating_fin'])
    context['analisi_current_ratio'] = get_analisi_current_ratio(context['rating_fin'])
    context['analisi_quick_ratio'] = get_analisi_quick_ratio(context['rating_fin'])
    context['analisi_posizionamento_fin'] = get_analisi_posizionamento_fin(context['rating_fin'])
    
    # Calcolo dei totali per la fascia di Rating dell'Azienda
    target_glob = context['rating_comb']
    
    if target_glob in ['A', 'B', 'C'] and 'Benchmark Totale' in df_rating.columns:
        num_soc_fascia_tot = len(df_rating[df_rating['Benchmark Totale'] == target_glob])
        tot_valide_bench = len(df_rating[df_rating['Benchmark Totale'].isin(['A', 'B', 'C'])])
        perc_soc_fascia_tot = (num_soc_fascia_tot / tot_valide_bench * 100) if tot_valide_bench > 0 else 0
        
        num_eco_fascia = len(df_rating[df_rating['Rating Economico'] == target_glob])
        num_patr_fascia = len(df_rating[df_rating['Rating Patrimoniale'] == target_glob])
        num_fin_fascia = len(df_rating[df_rating['Rating Finanziario'] == target_glob])
    else:
        num_soc_fascia_tot, perc_soc_fascia_tot, num_eco_fascia, num_patr_fascia, num_fin_fascia = 0, 0, 0, 0, 0

    context['analisi_combinata'] = get_analisi_combinata(context['rating_eco'], context['rating_patr'], context['rating_fin'])
    context['descr_fascia_appartenenza'] = get_descr_fascia_appartenenza(context['rating_comb'])
    context['num_soc_fascia_tot'] = f"{num_soc_fascia_tot:,}".replace(',', '.')
    context['perc_soc_fascia_tot'] = format_euro(perc_soc_fascia_tot)
    context['num_eco_fascia'] = f"{num_eco_fascia:,}".replace(',', '.')
    context['num_patr_fascia'] = f"{num_patr_fascia:,}".replace(',', '.')
    context['num_fin_fascia'] = f"{num_fin_fascia:,}".replace(',', '.')
    context['intro_divario_strutturale'] = get_intro_divario_strutturale(context['rating_eco'])
    context['analisi_trend_ebitda'] = get_analisi_trend_ebitda(context['rating_eco'], context['mg_ebitda'])
    context['asimmetria_ebitda'] = get_asimmetria_ebitda(context['rating_eco'])
    context['confronto_costi_settore'] = get_confronto_costi_settore(context['rating_eco'])
    context['impatto_costi_su_margine'] = get_impatto_costi_su_margine(context['rating_eco'])
    context['analisi_trend_ebit'] = get_analisi_trend_ebit(context['rating_eco'])
    context['confronto_ebit_settore'] = get_confronto_ebit_settore(context['rating_eco'])
    context['sintesi_ebit_mediana'] = get_sintesi_ebit_mediana(context['rating_eco'])
    context['analisi_verticale_ebit'] = get_analisi_verticale_ebit(context['rating_eco'])
    context['efficienza_struttura_asset'] = get_efficienza_struttura_asset(context['rating_eco'])
    context['analisi_trend_profitto'] = get_analisi_trend_profitto(context['rating_eco'])
    context['confronto_profitto_settore'] = get_confronto_profitto_settore(context['rating_eco'])
    context['prospettiva_redditivita_futura'] = get_prospettiva_redditivita_futura(context['rating_eco'])
    context['posizionamento_margine_profitto_fine'] = get_posizionamento_margine_profitto_fine(context['rating_eco'])
    context['sintesi_bilancio_finale'] = get_sintesi_bilancio_finale(context['rating_eco'])
    context['impatto_gestione_caratteristica'] = get_impatto_gestione_caratteristica(context['rating_eco'])
    context['impatto_oneri_accessori'] = get_impatto_oneri_accessori(context['rating_eco'])
    context['interpretazione_risultato_eco'] = get_interpretazione_risultato_eco(context['rating_eco'])
    context['stato_struttura_operativa'] = get_stato_struttura_operativa(context['rating_eco'])
    context['gestione_costi_ricavi_core'] = get_gestione_costi_ricavi_core(context['rating_eco'])
    context['implicazione_finale_eco'] = get_implicazione_finale_eco(context['rating_eco'])
    context['sintesi_quadriennio_patr'] = get_sintesi_quadriennio_patr(context['rating_patr'])
    context['analisi_copertura_attivo_fisso'] = get_analisi_copertura_attivo_fisso(context['rating_patr'])
    context['confronto_struttura_settore'] = get_confronto_struttura_settore(context['rating_patr'])
    context['analisi_soglia_struttura1'] = get_analisi_soglia_struttura1(context['rating_patr'])
    context['confronto_mediana_struttura1'] = get_confronto_mediana_struttura1(context['rating_patr'])
    context['implicazione_copertura_attivo'] = get_implicazione_copertura_attivo(context['rating_patr'])
    context['analisi_trend_struttura2'] = get_analisi_trend_struttura2(context['rating_patr'])
    context['confronto_settore_struttura2'] = get_confronto_settore_struttura2(context['rating_patr'])
    context['conclusione_struttura2'] = get_conclusione_struttura2(context['rating_patr'])
    context['evoluzione_vantaggio_competitivo'] = get_evoluzione_vantaggio_competitivo(context['rating_patr'])
    context['resilienza_credit_crunch'] = get_resilienza_credit_crunch(context['rating_patr'])
    context['flessibilita_strategica_asset'] = get_flessibilita_strategica_asset(context['rating_patr'])
    context['intro_analisi_gearing'] = get_intro_analisi_gearing(context['rating_patr'])
    context['andamento_storico_gearing'] = get_andamento_storico_gearing(context['rating_patr'])
    context['confronto_gearing_settore'] = get_confronto_gearing_settore(context['rating_patr'])
    context['reazione_contesto_gearing'] = get_reazione_contesto_gearing(context['rating_patr'])
    context['sintesi_valore_gearing'] = get_sintesi_valore_gearing(context['rating_patr'])
    context['analisi_rischio_default'] = get_analisi_rischio_default(context['rating_patr'])
    context['conclusione_autonomia_finanziaria'] = get_conclusione_autonomia_finanziaria(context['rating_patr'])
    context['sintesi_finale_patr'] = get_sintesi_finale_patr(context['rating_patr'])
    context['impatto_rischio_sistemico'] = get_impatto_rischio_sistemico(context['rating_patr'])
    context['posizionamento_competitivo_patr'] = get_posizionamento_competitivo_patr(context['rating_patr'])
    context['sintesi_modello_operativo_fin'] = get_sintesi_modello_operativo_fin(context['rating_fin'])
    context['analisi_trend_current_ratio'] = get_analisi_trend_current_ratio(context['rating_fin'])
    context['reazione_contesto_liquidita'] = get_reazione_contesto_liquidita(context['rating_fin'])
    context['confronto_current_ratio_mediana'] = get_confronto_current_ratio_mediana(context['rating_fin'])
    context['rapporto_attivita_passivita_brevi'] = get_rapporto_attivita_passivita_brevi(context['rating_fin'])
    context['capacita_generazione_liquidita'] = get_capacita_generazione_liquidita(context['rating_fin'])
    context['analisi_quick_ratio_soglia'] = get_analisi_quick_ratio_soglia(context['rating_fin'])
    context['implicazione_liquidita_immediata'] = get_implicazione_liquidita_immediata(context['rating_fin'])
    context['andamento_liquidita_primaria'] = get_andamento_liquidita_primaria(context['rating_fin'])
    context['confronto_quick_ratio_settore'] = get_confronto_quick_ratio_settore(context['rating_fin'])
    context['copertura_debiti_breve_quick'] = get_copertura_debiti_breve_quick(context['rating_fin'])
    context['interpretazione_flussi_cassa_quick'] = get_interpretazione_flussi_cassa_quick(context['rating_fin'])
    context['confronto_rotazione_mediana'] = get_confronto_rotazione_mediana(context['rating_fin'])
    context['interpretazione_modello_rotazione'] = get_interpretazione_modello_rotazione(context['rating_fin'])
    context['analisi_anomalia_rotazione'] = get_analisi_anomalia_rotazione(context['rating_fin'])
    context['rapporto_capitale_fatturato'] = get_rapporto_capitale_fatturato(context['rating_fin'])
    context['motivazione_strutturale_rotazione'] = get_motivazione_strutturale_rotazione(context['rating_fin'])
    context['sintesi_finale_fin'] = get_sintesi_finale_fin(context['rating_fin'])
    context['motivazione_rating_fin'] = get_motivazione_rating_fin(context['rating_fin'])
    context['gestione_tesoreria_fin'] = get_gestione_tesoreria_fin(context['rating_fin'])
    context['priorita_strategica_fin'] = get_priorita_strategica_fin(context['rating_fin'])
    context['sintesi_profilo_integrato'] = get_sintesi_profilo_integrato(context['rating_comb'])
    context['sintesi_posizionamento_lungo_periodo'] = get_sintesi_posizionamento_lungo_periodo(context['rating_comb'])
    context['conclusione_patrimoniale'] = get_conclusione_patrimoniale(context['rating_patr'])
    context['conclusione_economica'] = get_conclusione_economica(context['rating_eco'])
    context['conclusione_finanziaria_dettaglio'] = get_conclusione_finanziaria_dettaglio(context['rating_fin'])
    context['raccomandazione_finale'] = get_raccomandazione_finale(context['rating_comb'])

    context['impatto_territoriale'] = get_impatto_territoriale(perc_ricavi_target_su_macro, ragione_sociale_pulita, format_euro(ricavi_mln))

    # Costruzione delle Medaglie di Settore!!!
    conteggi = df_rating['Benchmark Totale'].value_counts()
    
    # =================================================================
    # 📚 NOTA METODOLOGICA E DIZIONARIO ISTAT
    # =================================================================
    
    # 1. Società valide (post-filtri) e Iniziali (dal foglio Risultati passato da finhack)
    totale_valide = len(df_rating)
    
    context['num_soc_valide'] = f"{len(df_rating):,}".replace(',', '.')
    context['num_max_soc'] = str(num_max_soc_orbis) # <--- Prende il 25.232 da Finhack!

    # 2. Dizionario ISTAT interno (tutti i settori)
    istat_db = {
        "00.10": 4751988,
        "B": 1881,
        "06": 16,
        "06.1": 6,
        "06.10": 6,
        "06.2": 10,
        "06.20": 10,
        "07": 2,
        "07.1": 1,
        "07.10": 1,
        "07.2": 1,
        "07.29": 1,
        "08": 1789,
        "08.1": 1627,
        "08.11": 768,
        "08.12": 859,
        "08.9": 162,
        "08.91": 4,
        "08.92": 3,
        "08.93": 22,
        "08.99": 133,
        "09": 74,
        "09.1": 44,
        "09.10": 44,
        "09.9": 30,
        "09.90": 30,
        "C": 355908,
        "10": 48051,
        "10.1": 3270,
        "10.11": 1353,
        "10.12": 108,
        "10.13": 1809,
        "10.2": 428,
        "10.20": 428,
        "10.3": 1714,
        "10.31": 44,
        "10.32": 127,
        "10.39": 1543,
        "10.4": 2780,
        "10.41": 2773,
        "10.42": 7,
        "10.5": 3084,
        "10.51": 2727,
        "10.52": 357,
        "10.6": 902,
        "10.61": 894,
        "10.62": 8,
        "10.7": 30951,
        "10.71": 25954,
        "10.72": 1346,
        "10.73": 3651,
        "10.8": 4411,
        "10.81": 11,
        "10.82": 564,
        "10.83": 897,
        "10.84": 294,
        "10.85": 987,
        "10.86": 253,
        "10.89": 1405,
        "10.9": 511,
        "10.91": 374,
        "10.92": 137,
        "11": 3203,
        "11.0": 3203,
        "11.01": 659,
        "11.02": 1582,
        "11.03": 5,
        "11.04": 66,
        "11.05": 684,
        "11.06": 3,
        "11.07": 204,
        "12": 10,
        "12.0": 10,
        "12.00": 10,
        "13": 10660,
        "13.1": 1153,
        "13.10": 1153,
        "13.2": 1328,
        "13.20": 1328,
        "13.3": 1787,
        "13.30": 1787,
        "13.9": 6392,
        "13.91": 682,
        "13.92": 3035,
        "13.93": 123,
        "13.94": 139,
        "13.95": 249,
        "13.96": 1073,
        "13.99": 1091,
        "14": 27641,
        "14.1": 25001,
        "14.11": 592,
        "14.12": 426,
        "14.13": 16853,
        "14.14": 880,
        "14.19": 6250,
        "14.2": 463,
        "14.20": 463,
        "14.3": 2177,
        "14.31": 434,
        "14.39": 1743,
        "15": 12408,
        "15.1": 6171,
        "15.11": 1575,
        "15.12": 4596,
        "15.2": 6237,
        "15.20": 6237,
        "16": 19873,
        "16.1": 2088,
        "16.10": 2088,
        "16.2": 17785,
        "16.21": 339,
        "16.22": 126,
        "16.23": 11838,
        "16.24": 1276,
        "16.29": 4206,
        "17": 3248,
        "17.1": 224,
        "17.11": 3,
        "17.12": 221,
        "17.2": 3024,
        "17.21": 1246,
        "17.22": 188,
        "17.23": 998,
        "17.24": 14,
        "17.29": 578,
        "18": 12375,
        "18.1": 12251,
        "18.11": 25,
        "18.12": 8968,
        "18.13": 2265,
        "18.14": 993,
        "18.2": 124,
        "18.20": 124,
        "19": 259,
        "19.1": 2,
        "19.10": 2,
        "19.2": 257,
        "19.20": 257,
        "20": 4242,
        "20.1": 925,
        "20.11": 81,
        "20.12": 81,
        "20.13": 133,
        "20.14": 90,
        "20.15": 184,
        "20.16": 344,
        "20.17": 12,
        "20.2": 33,
        "20.20": 33,
        "20.3": 721,
        "20.30": 721,
        "20.4": 1537,
        "20.41": 434,
        "20.42": 1103,
        "20.5": 1007,
        "20.51": 106,
        "20.52": 81,
        "20.53": 99,
        "20.59": 721,
        "20.6": 19,
        "20.60": 19,
        "21": 489,
        "21.1": 93,
        "21.10": 93,
        "21.2": 396,
        "21.20": 396,
        "22": 9109,
        "22.1": 1348,
        "22.19": 1261,
        "22.2": 7761,
        "22.21": 795,
        "22.22": 1452,
        "22.23": 745,
        "22.29": 4769,
        "23": 16150,
        "23.1": 2922,
        "23.11": 8,
        "23.12": 2142,
        "23.13": 53,
        "23.14": 32,
        "23.19": 687,
        "23.2": 87,
        "23.20": 87,
        "23.3": 494,
        "23.31": 260,
        "23.32": 234,
        "23.4": 2144,
        "23.41": 1854,
        "23.42": 61,
        "23.43": 6,
        "23.44": 36,
        "23.49": 187,
        "23.5": 107,
        "23.51": 46,
        "23.52": 61,
        "23.6": 2541,
        "23.61": 960,
        "23.62": 73,
        "23.63": 871,
        "23.64": 63,
        "23.65": 30,
        "23.69": 544,
        "23.7": 7055,
        "23.70": 7055,
        "23.9": 800,
        "23.91": 207,
        "23.99": 593,
        "24": 2948,
        "24.1": 370,
        "24.10": 370,
        "24.2": 350,
        "24.20": 350,
        "24.3": 817,
        "24.31": 14,
        "24.32": 37,
        "24.33": 622,
        "24.34": 144,
        "24.4": 572,
        "24.41": 163,
        "24.42": 249,
        "24.43": 28,
        "24.44": 57,
        "24.45": 75,
        "24.5": 839,
        "24.51": 117,
        "24.52": 31,
        "24.53": 437,
        "24.54": 254,
        "25": 71302,
        "25.1": 28486,
        "25.11": 12415,
        "25.12": 16071,
        "25.2": 535,
        "25.21": 97,
        "25.29": 438,
        "25.3": 91,
        "25.30": 91,
        "25.4": 114,
        "25.40": 114,
        "25.5": 1478,
        "25.50": 1478,
        "25.6": 23176,
        "25.61": 4506,
        "25.62": 18670,
        "25.7": 3877,
        "25.71": 303,
        "25.72": 523,
        "25.73": 3051,
        "25.9": 13545,
        "25.91": 74,
        "25.92": 149,
        "25.93": 585,
        "25.94": 321,
        "25.99": 12416,
        "26": 4818,
        "26.1": 1799,
        "26.11": 880,
        "26.12": 919,
        "26.2": 534,
        "26.20": 534,
        "26.3": 577,
        "26.30": 577,
        "26.4": 232,
        "26.40": 232,
        "26.5": 905,
        "26.51": 832,
        "26.52": 73,
        "26.6": 584,
        "26.60": 584,
        "26.7": 186,
        "26.70": 186,
        "26.8": 1,
        "26.80": 1,
        "27": 7108,
        "27.1": 1944,
        "27.11": 1010,
        "27.12": 934,
        "27.2": 79,
        "27.20": 79,
        "27.3": 780,
        "27.31": 14,
        "27.32": 402,
        "27.33": 364,
        "27.4": 1184,
        "27.40": 1184,
        "27.5": 473,
        "27.51": 329,
        "27.52": 144,
        "27.9": 2648,
        "27.90": 2648,
        "28": 18161,
        "28.1": 2232,
        "28.11": 172,
        "28.12": 406,
        "28.13": 400,
        "28.14": 792,
        "28.15": 462,
        "28.2": 7130,
        "28.21": 401,
        "28.22": 1107,
        "28.23": 128,
        "28.24": 33,
        "28.25": 992,
        "28.29": 4469,
        "28.3": 1322,
        "28.30": 1322,
        "28.4": 1664,
        "28.41": 627,
        "28.49": 1037,
        "28.9": 5813,
        "28.91": 567,
        "28.92": 420,
        "28.94": 805,
        "28.95": 285,
        "28.96": 479,
        "28.99": 1626,
        "29": 2451,
        "29.1": 143,
        "29.10": 143,
        "29.2": 826,
        "29.20": 826,
        "29.3": 1482,
        "29.31": 247,
        "29.32": 1235,
        "30": 3070,
        "30.1": 1953,
        "30.11": 616,
        "30.12": 1337,
        "30.2": 95,
        "30.20": 95,
        "30.3": 198,
        "30.30": 198,
        "30.4": 1,
        "30.40": 1,
        "30.9": 823,
        "30.91": 374,
        "30.92": 435,
        "30.99": 14,
        "31": 16439,
        "31.0": 16439,
        "31.01": 2102,
        "31.02": 619,
        "31.03": 517,
        "31.09": 13201,
        "32": 26718,
        "32.1": 7144,
        "32.11": 19,
        "32.12": 5534,
        "32.13": 1591,
        "32.2": 821,
        "32.20": 821,
        "32.3": 462,
        "32.30": 462,
        "32.4": 333,
        "32.40": 333,
        "32.5": 14780,
        "32.50": 14780,
        "32.9": 3178,
        "32.91": 179,
        "32.99": 2999,
        "33": 35175,
        "33.1": 27448,
        "33.11": 1761,
        "33.12": 17760,
        "33.13": 1710,
        "33.14": 1281,
        "33.15": 3399,
        "33.16": 138,
        "33.17": 188,
        "33.19": 1211,
        "33.2": 7727,
        "33.20": 7727,
        "D": 13269,
        "35": 13269,
        "35.1": 12485,
        "35.11": 10845,
        "35.12": 9,
        "35.13": 175,
        "35.14": 1456,
        "35.2": 550,
        "35.21": 63,
        "35.22": 179,
        "35.23": 308,
        "35.3": 234,
        "35.30": 234,
        "E": 9120,
        "36": 750,
        "36.0": 750,
        "36.00": 750,
        "37": 1474,
        "37.0": 1474,
        "37.00": 1474,
        "38": 6211,
        "38.1": 1935,
        "38.11": 1624,
        "38.12": 311,
        "38.2": 1076,
        "38.21": 931,
        "38.22": 145,
        "38.3": 3200,
        "38.31": 484,
        "38.32": 2716,
        "39": 685,
        "39.0": 685,
        "39.00": 685,
        "F": 544886,
        "41": 145938,
        "41.1": 5013,
        "41.10": 5013,
        "41.2": 140925,
        "41.20": 140925,
        "42": 8020,
        "42.1": 4105,
        "42.11": 3832,
        "42.12": 205,
        "42.13": 68,
        "42.2": 774,
        "42.21": 336,
        "42.22": 438,
        "42.9": 3141,
        "42.91": 344,
        "42.99": 2797,
        "43": 390928,
        "43.1": 10681,
        "43.11": 1371,
        "43.12": 8176,
        "43.13": 1134,
        "43.2": 144130,
        "43.21": 67851,
        "43.22": 62600,
        "43.29": 13679,
        "43.3": 219632,
        "43.31": 7258,
        "43.32": 22770,
        "43.33": 20575,
        "43.34": 35337,
        "43.39": 133692,
        "43.9": 16485,
        "43.91": 5173,
        "43.99": 11312,
        "45": 121034,
        "45.1": 27891,
        "45.11": 26204,
        "45.19": 1687,
        "45.2": 74920,
        "45.20": 74920,
        "45.3": 11010,
        "45.31": 5556,
        "45.32": 5454,
        "45.4": 7213,
        "45.40": 7213,
        "46": 345876,
        "46.1": 191132,
        "46.11": 3362,
        "46.12": 5291,
        "46.13": 10104,
        "46.14": 8886,
        "46.15": 11424,
        "46.16": 10446,
        "46.17": 32307,
        "46.18": 47873,
        "46.19": 61439,
        "46.2": 6969,
        "46.22": 1441,
        "46.23": 1180,
        "46.24": 1109,
        "46.31": 7503,
        "46.32": 2357,
        "46.34": 5533,
        "46.35": 68,
        "46.4": 43425,
        "46.41": 2710,
        "46.42": 9693,
        "46.45": 3284,
        "46.46": 5664,
        "46.48": 2108,
        "46.49": 10714,
        "46.5": 6730,
        "46.52": 2117,
        "46.62": 2544,
        "46.65": 890,
        "46.66": 1593,
        "46.69": 10342,
        "46.7": 37210,
        "46.72": 3251,
        "46.75": 3186,
        "46.76": 3372,
        "46.77": 4174,
        "46.9": 8368,
        "46.90": 8368,
        "47": 517833,
        "47.1": 55518,
        "47.11": 43948,
        "47.19": 11570,
        "47.2": 90280,
        "47.21": 13365,
        "47.22": 19637,
        "47.23": 4957,
        "47.24": 4454,
        "47.25": 4263,
        "47.26": 30978,
        "47.29": 12626,
        "47.3": 12428,
        "47.30": 12428,
        "47.4": 10547,
        "47.41": 4977,
        "47.42": 4968,
        "47.43": 602,
        "47.5": 60003,
        "47.51": 8265,
        "47.52": 25009,
        "47.53": 1301,
        "47.54": 2596,
        "47.59": 22832,
        "47.6": 29715,
        "47.61": 3117,
        "47.62": 15817,
        "47.63": 484,
        "47.64": 7481,
        "47.65": 2816,
        "47.7": 168226,
        "47.71": 54535,
        "47.72": 10888,
        "47.73": 20210,
        "47.74": 3285,
        "47.75": 10625,
        "47.76": 15562,
        "47.77": 12299,
        "47.78": 36556,
        "47.79": 4266,
        "47.8": 63459,
        "47.81": 19893,
        "47.82": 24459,
        "47.89": 19107,
        "47.9": 27657,
        "47.91": 20418,
        "47.99": 7239,
        "H": 120213,
        "49": 90562,
        "49.1": 17,
        "49.10": 17,
        "49.2": 22,
        "49.20": 22,
        "49.3": 32255,
        "49.31": 911,
        "49.32": 27996,
        "49.39": 3348,
        "49.4": 58252,
        "49.41": 56836,
        "49.42": 1416,
        "49.5": 16,
        "49.50": 16,
        "50": 2586,
        "50.1": 1104,
        "50.10": 1104,
        "50.2": 149,
        "50.20": 149,
        "50.3": 1226,
        "50.30": 1226,
        "50.4": 107,
        "50.40": 107,
        "51": 131,
        "51.1": 112,
        "51.10": 112,
        "51.2": 19,
        "51.21": 19,
        "52": 21781,
        "52.1": 1489,
        "52.10": 1489,
        "52.2": 20292,
        "52.21": 6991,
        "52.22": 2005,
        "52.23": 430,
        "52.24": 3065,
        "52.29": 7801,
        "53": 5153,
        "53.1": 1,
        "53.10": 1,
        "53.2": 5152,
        "53.20": 5152,
        "I": 338404,
        "55": 68529,
        "55.1": 22729,
        "55.10": 22729,
        "55.2": 43935,
        "55.20": 43935,
        "55.3": 1618,
        "55.30": 1618,
        "55.9": 247,
        "55.90": 247,
        "56": 269875,
        "56.1": 165258,
        "56.10": 165258,
        "56.2": 3672,
        "56.21": 2222,
        "56.29": 1450,
        "56.3": 100945,
        "56.30": 100945,
        "J": 125648,
        "58": 5726,
        "58.1": 4886,
        "58.11": 1929,
        "58.12": 7,
        "58.13": 401,
        "58.14": 1423,
        "58.19": 1126,
        "58.2": 840,
        "58.21": 22,
        "58.29": 818,
        "59": 10716,
        "59.1": 8798,
        "59.11": 6134,
        "59.12": 1800,
        "59.13": 220,
        "59.14": 644,
        "59.2": 1918,
        "59.20": 1918,
        "60": 1267,
        "60.1": 702,
        "60.10": 702,
        "60.2": 565,
        "60.20": 565,
        "61": 3512,
        "61.1": 264,
        "61.10": 264,
        "61.2": 119,
        "61.20": 119,
        "61.3": 24,
        "61.30": 24,
        "61.9": 3105,
        "61.90": 3105,
        "62": 63327,
        "62.0": 63327,
        "62.01": 20750,
        "62.02": 29753,
        "62.03": 1063,
        "62.09": 11761,
        "63": 41100,
        "63.1": 37562,
        "63.11": 35135,
        "63.12": 2427,
        "63.9": 3538,
        "63.91": 509,
        "63.99": 3029,
        "K": 120248,
        "64": 23826,
        "64.1": 389,
        "64.11": 1,
        "64.19": 388,
        "64.2": 20522,
        "64.20": 20522,
        "64.3": 79,
        "64.30": 79,
        "64.9": 2836,
        "64.91": 144,
        "64.92": 477,
        "64.99": 2215,
        "65": 182,
        "65.1": 175,
        "65.11": 59,
        "65.12": 116,
        "65.2": 7,
        "65.20": 7,
        "66": 96240,
        "66.1": 39609,
        "66.11": 5,
        "66.12": 103,
        "66.19": 39501,
        "66.2": 56354,
        "66.21": 5097,
        "66.22": 50779,
        "66.29": 478,
        "66.3": 277,
        "66.30": 277,
        "L": 250621,
        "68": 250621,
        "68.1": 29568,
        "68.10": 29568,
        "68.2": 159486,
        "68.20": 159486,
        "68.3": 61567,
        "68.31": 42633,
        "68.32": 18934,
        "M": 934227,
        "69": 319307,
        "69.1": 182798,
        "69.10": 182798,
        "69.2": 136509,
        "69.20": 136509,
        "70": 117805,
        "70.1": 2544,
        "70.10": 2544,
        "70.2": 115261,
        "70.21": 9207,
        "70.22": 106054,
        "71.1": 237345,
        "71.11": 82790,
        "71.12": 154555,
        "71.2": 8385,
        "71.20": 8385,
        "72": 16440,
        "72.1": 14471,
        "72.11": 8369,
        "72.19": 6102,
        "72.2": 1969,
        "72.20": 1969,
        "73": 36336,
        "73.1": 30972,
        "73.11": 28986,
        "73.12": 1986,
        "73.2": 5364,
        "73.20": 5364,
        "74": 179867,
        "74.1": 47980,
        "74.10": 47980,
        "74.2": 19783,
        "74.20": 19783,
        "74.3": 8228,
        "74.30": 8228,
        "74.9": 103876,
        "74.90": 103876,
        "75": 18742,
        "75.0": 18742,
        "75.00": 18742,
        "N": 195326,
        "77": 17973,
        "77.1": 4677,
        "77.11": 4337,
        "77.12": 340,
        "77.2": 4202,
        "77.21": 3362,
        "77.22": 70,
        "77.29": 770,
        "77.3": 7969,
        "77.31": 123,
        "77.32": 1601,
        "77.33": 557,
        "77.34": 755,
        "77.35": 122,
        "77.39": 4811,
        "77.4": 1125,
        "77.40": 1125,
        "78": 1137,
        "78.1": 996,
        "78.10": 996,
        "78.2": 133,
        "78.20": 133,
        "78.3": 8,
        "78.30": 8,
        "79": 22215,
        "79.1": 9132,
        "79.11": 7451,
        "79.12": 1681,
        "79.9": 13083,
        "79.90": 13083,
        "80": 2415,
        "80.1": 1238,
        "80.10": 1238,
        "80.2": 209,
        "80.20": 209,
        "80.3": 968,
        "80.30": 968,
        "81": 64408,
        "81.1": 2853,
        "81.10": 2853,
        "81.2": 40117,
        "81.21": 32882,
        "81.22": 1845,
        "81.29": 5390,
        "81.3": 21438,
        "81.30": 21438,
        "82": 87178,
        "82.1": 12964,
        "82.11": 9054,
        "82.19": 3910,
        "82.2": 1360,
        "82.20": 1360,
        "82.3": 4892,
        "82.30": 4892,
        "82.9": 67962,
        "82.91": 1418,
        "82.92": 2258,
        "82.99": 64286,
        "P": 54890,
        "85": 54890,
        "85.1": 1468,
        "85.10": 1468,
        "85.2": 269,
        "85.20": 269,
        "85.3": 1405,
        "85.31": 344,
        "85.32": 1061,
        "85.4": 855,
        "85.41": 210,
        "85.42": 645,
        "85.5": 48632,
        "85.51": 15321,
        "85.52": 5543,
        "85.53": 4591,
        "85.59": 23177,
        "85.6": 2261,
        "85.60": 2261,
        "Q": 384369,
        "86": 364376,
        "86.1": 877,
        "86.10": 877,
        "86.2": 210234,
        "86.21": 65517,
        "86.22": 89820,
        "86.23": 54897,
        "86.9": 153265,
        "86.90": 153265,
        "87": 6777,
        "87.1": 930,
        "87.10": 930,
        "87.2": 573,
        "87.20": 573,
        "87.3": 3593,
        "87.30": 3593,
        "87.9": 1681,
        "87.90": 1681,
        "88": 13216,
        "88.1": 2941,
        "88.10": 2941,
        "88.9": 10275,
        "88.91": 4063,
        "88.99": 6212,
        "R": 91859,
        "90": 45822,
        "90.0": 45822,
        "90.01": 15585,
        "90.02": 9439,
        "90.03": 20516,
        "90.04": 282,
        "91": 1228,
        "91.0": 1228,
        "91.01": 594,
        "91.02": 304,
        "91.03": 230,
        "91.04": 100,
        "92": 7305,
        "92.0": 7305,
        "92.00": 7305,
        "93": 37504,
        "93.1": 20843,
        "93.11": 3894,
        "93.12": 2471,
        "93.13": 4470,
        "93.19": 10008,
        "93.2": 16661,
        "93.21": 1775,
        "93.29": 14886,
        "S": 226376,
        "95": 24924,
        "95.1": 5795,
        "95.11": 3378,
        "95.12": 2417,
        "95.2": 19129,
        "95.21": 959,
        "95.22": 3150,
        "95.23": 2324,
        "95.24": 5705,
        "95.25": 1571,
        "95.29": 5420,
        "96": 201452,
        "96.0": 201452,
        "96.01": 14453,
        "96.02": 141437,
        "96.03": 6857,
        "96.04": 2946,
        "96.09": 35759,
    }

    totale_istat = istat_db.get(cod_nace_pulito, 0)
    
    if totale_istat > 0:
        context['max_soc_istat'] = f"{totale_istat:,}".replace(',', '.')
    else:
        context['max_soc_istat'] = "N/D"

    # 3. Calcolo percentuale automatica
    try:
        if totale_istat > 0:
            percentuale = (totale_valide / totale_istat) * 100
            context['perc_su_istat'] = f"{percentuale:.2f}".replace('.', ',')
        else:
            context['perc_su_istat'] = "N/D"
    except (ValueError, TypeError, ZeroDivisionError):
        context['perc_su_istat'] = "N/D"

    # --- 

    if len(conteggi) > 0:
        rat1_nome = str(conteggi.index[0]).strip()
        rat1_num = conteggi.iloc[0]
        context['rating_piu_presente'] = rat1_nome
        context['rating_piu_pres'] = rat1_nome
        context['rat1_piu_pres_categ'] = rat1_nome
        context['rat1_piu_pres_num'] = f"{rat1_num:,}".replace(',', '.')
        context['rating_piu_pres_num_tot'] = f"{rat1_num:,}".replace(',', '.')

        if 'A' in rat1_nome: desc = 'Positivo / Solido'
        elif 'B' in rat1_nome: desc = 'Medio / Equilibrato'
        else: desc = 'Rischioso / Critico'
        context['assoc_desc_rating_lett'] = desc

    if len(conteggi) > 1:
        context['rat2_piu_pres_categ'] = str(conteggi.index[1]).strip()
        context['rat2_piu_pres_num'] = f"{conteggi.iloc[1]:,}".replace(',', '.')

    if len(conteggi) > 2:
        context['rat3_piu_pres_categ'] = str(conteggi.index[2]).strip()
        context['rat3_piu_pres_num'] = f"{conteggi.iloc[2]:,}".replace(',', '.')
    
    # =================================================================
    # VINCITORI DELLE SINGOLE AREE (Per il paragrafo descrittivo)
    # =================================================================
    
    # Economico
    eco_counts = df_rating['Rating Economico'].value_counts()
    if len(eco_counts) > 0:
        context['rat_eco_piu_pres'] = str(eco_counts.index[0]).strip()
        context['num_eco_piu_pres'] = f"{eco_counts.iloc[0]:,}".replace(',', '.')
    else:
        context['rat_eco_piu_pres'], context['num_eco_piu_pres'] = "N.D.", "0"

    # Patrimoniale
    patr_counts = df_rating['Rating Patrimoniale'].value_counts()
    if len(patr_counts) > 0:
        context['rat_patr_piu_pres'] = str(patr_counts.index[0]).strip()
        context['num_patr_piu_pres'] = f"{patr_counts.iloc[0]:,}".replace(',', '.')
    else:
        context['rat_patr_piu_pres'], context['num_patr_piu_pres'] = "N.D.", "0"

    # Finanziario
    fin_counts = df_rating['Rating Finanziario'].value_counts()
    if len(fin_counts) > 0:
        context['rat_fin_piu_pres'] = str(fin_counts.index[0]).strip()
        context['num_fin_piu_pres'] = f"{fin_counts.iloc[0]:,}".replace(',', '.')
    else:
        context['rat_fin_piu_pres'], context['num_fin_piu_pres'] = "N.D.", "0"
    
    # =================================================================
    # FORMA DELLA DISTRIBUZIONE (Asimmetria e Curtosi sui Margini)
    # Utilizziamo l'EBITDA 2024 come proxy rappresentativo della redditività
    # =================================================================
    col_ebitda = 'Margine EBITDA (*) % 2024'
    if col_ebitda in df_orbis.columns:
        skew_val = df_orbis[col_ebitda].skew()
        kurt_val = df_orbis[col_ebitda].kurt() # In Pandas > 0 è leptocurtica

        # Analisi Asimmetria
        if skew_val > 0.3:
            context['tipo_asimmetria'] = "un'asimmetria positiva"
            context['rel_media_mediana'] = "dall'evidente scostamento tra la media (maggiore) e la mediana"
        elif skew_val < -0.3:
            context['tipo_asimmetria'] = "un'asimmetria negativa"
            context['rel_media_mediana'] = "dall'evidente scostamento tra la media (minore) e la mediana"
        else:
            context['tipo_asimmetria'] = "una sostanziale simmetria"
            context['rel_media_mediana'] = "dal generale allineamento tra i valori di media e mediana"

        # Analisi Curtosi
        if kurt_val > 0.5:
            context['tipo_curtosi'] = "leptocurtiche"
        elif kurt_val < -0.5:
            context['tipo_curtosi'] = "platicurtiche"
        else:
            context['tipo_curtosi'] = "mesocurtiche"
    else:
        context['tipo_asimmetria'] = "dati non sufficienti"
        context['rel_media_mediana'] = "N.D."
        context['tipo_curtosi'] = "non definibili"

    # =================================================================
    # FORMA DELLA DISTRIBUZIONE (Asimmetria e Curtosi Patrimoniale)
    # Utilizziamo l'Indice di Struttura 1° livello come proxy rappresentativo
    # =================================================================
    col_strut = 'Indice di Struttura 1° livello (*) 2024'
    if col_strut in df_orbis.columns:
        skew_patr = df_orbis[col_strut].skew()
        kurt_patr = df_orbis[col_strut].kurt()

        # Analisi Asimmetria Patrimoniale
        if skew_patr > 0.3:
            context['tipo_asimmetria_patr'] = "asimmetriche, in questo caso positive,"
        elif skew_patr < -0.3:
            context['tipo_asimmetria_patr'] = "asimmetriche, in questo caso negative,"
        else:
            context['tipo_asimmetria_patr'] = "sostanzialmente simmetriche"

        # Analisi Curtosi Patrimoniale
        if kurt_patr > 0.5:
            context['tipo_curtosi_patr'] = "leptocurtiche"
        elif kurt_patr < -0.5:
            context['tipo_curtosi_patr'] = "platicurtiche"
        else:
            context['tipo_curtosi_patr'] = "mesocurtiche"
    else:
        context['tipo_asimmetria_patr'] = "non valutabili"
        context['tipo_curtosi_patr'] = "non definibili"
    
    # =================================================================
    # FORMA DELLA DISTRIBUZIONE (Asimmetria e Curtosi Finanziaria)
    # Utilizziamo il Current Ratio come proxy rappresentativo
    # =================================================================
    col_fin = 'Current Ratio (*) 2024'
    if col_fin in df_orbis.columns:
        skew_fin = df_orbis[col_fin].skew()
        kurt_fin = df_orbis[col_fin].kurt()

        # Analisi Asimmetria Finanziaria
        if skew_fin > 0.3:
            context['tipo_asimmetria_fin'] = "un'asimmetria positiva"
        elif skew_fin < -0.3:
            context['tipo_asimmetria_fin'] = "un'asimmetria negativa"
        else:
            context['tipo_asimmetria_fin'] = "una sostanziale simmetria"

        # Analisi Curtosi Finanziaria
        if kurt_fin > 0.5:
            context['tipo_curtosi_fin'] = "leptocurtiche"
        elif kurt_fin < -0.5:
            context['tipo_curtosi_fin'] = "platicurtiche"
        else:
            context['tipo_curtosi_fin'] = "mesocurtiche"
    else:
        context['tipo_asimmetria_fin'] = "non valutabili"
        context['tipo_curtosi_fin'] = "non definibili"
    
    # ---

    # 🟢 ATTIVIAMO LA LAVATRICE NUCLEARE
    template_pulito = lavatrice_nucleare(template_path)
    doc = DocxTemplate(template_pulito)

    # =================================================================
    # COSTRUZIONE TABELLA 1 DINAMICA DA PYTHON
    # =================================================================
    sd_tab1 = doc.new_subdoc()
    t1 = sd_tab1.add_table(rows=1, cols=9)
    t1.style = 'Table Grid' # Stile classico a griglia di Word

    # 1. Creiamo le Intestazioni in Grassetto
    intestazioni = ['Regioni', 'Imprese V.A.', 'Imprese %', 'Ricavi V.A.', 'Ricavi %', 'Attivo V.A.', 'Attivo %', 'Dip. V.A.', 'Dip. %']
    for i, intestazione in enumerate(intestazioni):
        t1.rows[0].cells[i].text = intestazione
        t1.rows[0].cells[i].paragraphs[0].runs[0].bold = True

    # 2. Calcoli Territoriali
    df_terr = df_orbis.copy()
    def pulisci_regione(x):
        if pd.isna(x): return 'Altro'
        return str(x).split(' - ')[1] if ' - ' in str(x) else str(x)
    df_terr['Reg_Clean'] = df_terr[col_regione].apply(pulisci_regione) if col_regione else 'Altro'

    tot_imp = len(df_terr)
    tot_ric = df_terr['Totale valore della produzione migl EUR 2024'].sum()
    tot_att = df_terr['Totale Attivo migl EUR 2024'].sum()
    tot_dip = df_terr['Numero dipendenti 2024'].sum()

    pivot_terr = df_terr.groupby(['Macroregione', 'Reg_Clean']).agg({
        col_ragione: 'count', 'Totale valore della produzione migl EUR 2024': 'sum', 
        'Totale Attivo migl EUR 2024': 'sum', 'Numero dipendenti 2024': 'sum'
    }).reset_index()

    # 3. Compilazione Righe
    for macro in ['Nord Ovest', 'Nord Est', 'Centro', 'Sud e Isole', 'Altro']:
        df_m = pivot_terr[pivot_terr['Macroregione'] == macro]
        if df_m.empty: continue
        
       # Inserimento Singole Regioni
        for _, r in df_m.iterrows():
            row_cells = t1.add_row().cells
            nome_regione = str(r['Reg_Clean'])
            
            row_cells[0].text = nome_regione
            row_cells[1].text = f"{int(r[col_ragione]):,}".replace(',', '.')
            row_cells[2].text = format_euro((r[col_ragione]/tot_imp)*100) if tot_imp else "0,00"
            row_cells[3].text = format_euro(r['Totale valore della produzione migl EUR 2024'])
            row_cells[4].text = format_euro((r['Totale valore della produzione migl EUR 2024']/tot_ric)*100) if tot_ric else "0,00"
            row_cells[5].text = format_euro(r['Totale Attivo migl EUR 2024'])
            row_cells[6].text = format_euro((r['Totale Attivo migl EUR 2024']/tot_att)*100) if tot_att else "0,00"
            row_cells[7].text = f"{int(r['Numero dipendenti 2024']):,}".replace(',', '.')
            row_cells[8].text = format_euro((r['Numero dipendenti 2024']/tot_dip)*100) if tot_dip else "0,00"

            # 🎯 EVIDENZIA LA REGIONE DELL'AZIENDA (Sfondo azzurro e grassetto)
            if nome_regione.strip().lower() == regione_target_pulita.strip().lower():
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls
                for cell in row_cells:
                    shd = parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="DDEBF7"/>' % nsdecls('w'))
                    cell._tc.get_or_add_tcPr().append(shd)
                    if len(cell.paragraphs[0].runs) > 0:
                        cell.paragraphs[0].runs[0].bold = True
            
        # Inserimento Totale Macroregione (Grassetto)
        row_cells = t1.add_row().cells
        row_cells[0].text = f"TOTALE {macro.upper()}"
        row_cells[0].paragraphs[0].runs[0].bold = True
        row_cells[1].text = f"{int(df_m[col_ragione].sum()):,}".replace(',', '.')
        row_cells[2].text = format_euro((df_m[col_ragione].sum()/tot_imp)*100) if tot_imp else "0,00"
        row_cells[3].text = format_euro(df_m['Totale valore della produzione migl EUR 2024'].sum())
        row_cells[4].text = format_euro((df_m['Totale valore della produzione migl EUR 2024'].sum()/tot_ric)*100) if tot_ric else "0,00"
        row_cells[5].text = format_euro(df_m['Totale Attivo migl EUR 2024'].sum())
        row_cells[6].text = format_euro((df_m['Totale Attivo migl EUR 2024'].sum()/tot_att)*100) if tot_att else "0,00"
        row_cells[7].text = f"{int(df_m['Numero dipendenti 2024'].sum()):,}".replace(',', '.')
        row_cells[8].text = format_euro((df_m['Numero dipendenti 2024'].sum()/tot_dip)*100) if tot_dip else "0,00"

    # Inserimento Totale Italia Finale (Tutto in Grassetto)
    row_cells = t1.add_row().cells
    row_cells[0].text = "ITALIA (TOTALE)"
    row_cells[1].text = f"{int(tot_imp):,}".replace(',', '.')
    row_cells[2].text = "100,00"
    row_cells[3].text = format_euro(tot_ric)
    row_cells[4].text = "100,00"
    row_cells[5].text = format_euro(tot_att)
    row_cells[6].text = "100,00"
    row_cells[7].text = f"{int(tot_dip):,}".replace(',', '.')
    row_cells[8].text = "100,00"
    for cell in row_cells: cell.paragraphs[0].runs[0].bold = True

    for row in t1.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(10.5)

    # Iniettiamo la tabella nel documento
    context['tabella_1_dinamica'] = sd_tab1


    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 2: CLASSIFICAZIONE POSIZIONAMENTO
    # =================================================================
    sd_tab2 = doc.new_subdoc()
    t2 = sd_tab2.add_table(rows=3, cols=4)
    t2.style = 'Table Grid'

    # Riga 0: Ragione Sociale (Uniamo le 4 celle in una sola)
    cella_unita = t2.cell(0, 0)
    cella_unita.merge(t2.cell(0, 3))
    cella_unita.text = context.get('ragione_sociale', 'Azienda Target')
    cella_unita.paragraphs[0].runs[0].bold = True

    # Riga 1: Intestazioni
    headers_t2 = ['Benchmark Economico', 'Benchmark Patrimoniale', 'Benchmark Finanziario', 'Rating Combinato']
    for i, h in enumerate(headers_t2):
        t2.cell(1, i).text = h
        t2.cell(1, i).paragraphs[0].runs[0].bold = True

    # Riga 2: Valori (es. A, B, C, ABC)
    valori_t2 = [
        context.get('rating_eco', 'N.D.'), 
        context.get('rating_patr', 'N.D.'), 
        context.get('rating_fin', 'N.D.'), 
        context.get('rating_tot', 'N.D.')
    ]
    for i, v in enumerate(valori_t2):
        t2.cell(2, i).text = str(v)

    # Iniettiamo la tabella nel documento
    context['tabella_2_dinamica'] = sd_tab2

    # =================================================================
    # 🏗️ COSTRUZIONE TABELLE 3, 4, 5: RANKING AZIENDALI
    # =================================================================

    # --- TABELLA 3: ECONOMICO ---
    sd_tab3 = doc.new_subdoc()
    t3 = sd_tab3.add_table(rows=4, cols=4)
    t3.style = 'Table Grid'
    t3.cell(0, 1).merge(t3.cell(0, 3))
    t3.cell(0, 1).text = 'Equilibrio Economico'
    t3.cell(0, 1).paragraphs[0].runs[0].bold = True

    t3.cell(1, 1).text = "Margine EBITDA"
    t3.cell(1, 2).text = "Margine EBIT"
    t3.cell(1, 3).text = "Margine di Profitto"
    for i in range(1, 4):
        t3.cell(1, i).paragraphs[0].runs[0].bold = True

    t3.cell(2, 0).text = 'Ranking Nazionale'
    t3.cell(2, 0).paragraphs[0].runs[0].bold = True
    t3.cell(2, 1).text = f"{context.get('rnk_naz_ebitda', 'N.D.')}°"
    t3.cell(2, 2).text = f"{context.get('rnk_naz_ebit', 'N.D.')}°"
    t3.cell(2, 3).text = f"{context.get('rnk_naz_prof', 'N.D.')}°"

    t3.cell(3, 0).text = 'Ranking Regionale'
    t3.cell(3, 0).paragraphs[0].runs[0].bold = True
    t3.cell(3, 1).text = f"{context.get('rnk_reg_ebitda', 'N.D.')}°"
    t3.cell(3, 2).text = f"{context.get('rnk_reg_ebit', 'N.D.')}°"
    t3.cell(3, 3).text = f"{context.get('rnk_reg_prof', 'N.D.')}°"
    context['tabella_3_dinamica'] = sd_tab3

    # --- TABELLA 4: PATRIMONIALE ---
    sd_tab4 = doc.new_subdoc()
    t4 = sd_tab4.add_table(rows=4, cols=4)
    t4.style = 'Table Grid'
    t4.cell(0, 1).merge(t4.cell(0, 3))
    t4.cell(0, 1).text = 'Equilibrio Patrimoniale'
    t4.cell(0, 1).paragraphs[0].runs[0].bold = True

    t4.cell(1, 1).text = "Indice Struttura 1° Liv."
    t4.cell(1, 2).text = "Indice Struttura 2° Liv."
    t4.cell(1, 3).text = "Gearing"
    for i in range(1, 4):
        t4.cell(1, i).paragraphs[0].runs[0].bold = True

    t4.cell(2, 0).text = 'Ranking Nazionale'
    t4.cell(2, 0).paragraphs[0].runs[0].bold = True
    t4.cell(2, 1).text = f"{context.get('rnk_naz_strut1', 'N.D.')}°"
    t4.cell(2, 2).text = f"{context.get('rnk_naz_strut2', 'N.D.')}°"
    t4.cell(2, 3).text = f"{context.get('rnk_naz_gear', 'N.D.')}°"

    t4.cell(3, 0).text = 'Ranking Regionale'
    t4.cell(3, 0).paragraphs[0].runs[0].bold = True
    t4.cell(3, 1).text = f"{context.get('rnk_reg_strut1', 'N.D.')}°"
    t4.cell(3, 2).text = f"{context.get('rnk_reg_strut2', 'N.D.')}°"
    t4.cell(3, 3).text = f"{context.get('rnk_reg_gear', 'N.D.')}°"
    context['tabella_4_dinamica'] = sd_tab4

    # --- TABELLA 5: FINANZIARIO ---
    sd_tab5 = doc.new_subdoc()
    t5 = sd_tab5.add_table(rows=4, cols=4)
    t5.style = 'Table Grid'
    t5.cell(0, 1).merge(t5.cell(0, 3))
    t5.cell(0, 1).text = 'Equilibrio Finanziario'
    t5.cell(0, 1).paragraphs[0].runs[0].bold = True

    t5.cell(1, 1).text = "Current Ratio"
    t5.cell(1, 2).text = "Quick Ratio"
    t5.cell(1, 3).text = "Indice Rot.Cap.Inv."
    for i in range(1, 4):
        t5.cell(1, i).paragraphs[0].runs[0].bold = True

    t5.cell(2, 0).text = 'Ranking Nazionale'
    t5.cell(2, 0).paragraphs[0].runs[0].bold = True
    t5.cell(2, 1).text = f"{context.get('rnk_naz_cr', 'N.D.')}°"
    t5.cell(2, 2).text = f"{context.get('rnk_naz_qr', 'N.D.')}°"
    t5.cell(2, 3).text = f"{context.get('rnk_naz_rot', 'N.D.')}°"

    t5.cell(3, 0).text = 'Ranking Regionale'
    t5.cell(3, 0).paragraphs[0].runs[0].bold = True
    t5.cell(3, 1).text = f"{context.get('rnk_reg_cr', 'N.D.')}°"
    t5.cell(3, 2).text = f"{context.get('rnk_reg_qr', 'N.D.')}°"
    t5.cell(3, 3).text = f"{context.get('rnk_reg_rot', 'N.D.')}°"
    context['tabella_5_dinamica'] = sd_tab5


    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 6: RIPARTIZIONE TERRITORIALE BENCHMARK TOTALE
    # =================================================================
    sd_tab6 = doc.new_subdoc()
    t6 = sd_tab6.add_table(rows=2, cols=9)
    t6.style = 'Table Grid'

    # Uniamo le celle per le intestazioni della riga superiore
    t6.cell(0, 0).merge(t6.cell(1, 0))
    t6.cell(0, 0).text = 'Regioni'
    
    t6.cell(0, 1).merge(t6.cell(0, 2))
    t6.cell(0, 1).text = 'Totale Imprese'
    
    t6.cell(0, 3).merge(t6.cell(0, 4))
    t6.cell(0, 3).text = 'Classe A'
    
    t6.cell(0, 5).merge(t6.cell(0, 6))
    t6.cell(0, 5).text = 'Classe B'
    
    t6.cell(0, 7).merge(t6.cell(0, 8))
    t6.cell(0, 7).text = 'Classe C'

    # Sotto-intestazioni (V.A. e %)
    for col_idx in [1, 3, 5, 7]:
        t6.cell(1, col_idx).text = 'V.A.'
        t6.cell(1, col_idx+1).text = '%'

    # Formattiamo le intestazioni in grassetto
    for r_idx in [0, 1]:
        for c_idx in range(9):
            if len(t6.cell(r_idx, c_idx).paragraphs[0].runs) > 0:
                t6.cell(r_idx, c_idx).paragraphs[0].runs[0].bold = True

    # Calcoli per la Tabella 6
    df_t6 = df_rating.copy()
    def pulisci_reg(x):
        if pd.isna(x): return 'Altro'
        return str(x).split(' - ')[1] if ' - ' in str(x) else str(x)
    
    df_t6['Reg_Clean'] = df_t6[col_regione].apply(pulisci_reg) if col_regione else 'Altro'

    if 'Benchmark Totale' in df_t6.columns:
        tot_ita = len(df_t6)
        tot_ita_A = len(df_t6[df_t6['Benchmark Totale'] == 'A'])
        tot_ita_B = len(df_t6[df_t6['Benchmark Totale'] == 'B'])
        tot_ita_C = len(df_t6[df_t6['Benchmark Totale'] == 'C'])
        
        for macro in ['Nord Ovest', 'Nord Est', 'Centro', 'Sud e Isole', 'Altro']:
            df_m = df_t6[df_t6['Macroregione'] == macro]
            if df_m.empty: continue
            
            regioni_macro = df_m['Reg_Clean'].unique()
            for reg in sorted(regioni_macro):
                df_reg = df_m[df_m['Reg_Clean'] == reg]
                val_tot = len(df_reg)
                val_A = len(df_reg[df_reg['Benchmark Totale'] == 'A'])
                val_B = len(df_reg[df_reg['Benchmark Totale'] == 'B'])
                val_C = len(df_reg[df_reg['Benchmark Totale'] == 'C'])
                
                row_cells = t6.add_row().cells
                row_cells[0].text = str(reg)
                row_cells[1].text = f"{val_tot:,}".replace(',', '.')
                row_cells[2].text = format_euro((val_tot/tot_ita)*100) if tot_ita else "0,00"
                row_cells[3].text = f"{val_A:,}".replace(',', '.')
                row_cells[4].text = format_euro((val_A/tot_ita)*100) if tot_ita else "0,00"
                row_cells[5].text = f"{val_B:,}".replace(',', '.')
                row_cells[6].text = format_euro((val_B/tot_ita)*100) if tot_ita else "0,00"
                row_cells[7].text = f"{val_C:,}".replace(',', '.')
                row_cells[8].text = format_euro((val_C/tot_ita)*100) if tot_ita else "0,00"
                
            # Riga Totale Macroregione (Grassetto)
            val_tot_m = len(df_m)
            val_A_m = len(df_m[df_m['Benchmark Totale'] == 'A'])
            val_B_m = len(df_m[df_m['Benchmark Totale'] == 'B'])
            val_C_m = len(df_m[df_m['Benchmark Totale'] == 'C'])
            
            row_cells = t6.add_row().cells
            row_cells[0].text = f"TOTALE {macro.upper()}"
            row_cells[1].text = f"{val_tot_m:,}".replace(',', '.')
            row_cells[2].text = format_euro((val_tot_m/tot_ita)*100) if tot_ita else "0,00"
            row_cells[3].text = f"{val_A_m:,}".replace(',', '.')
            row_cells[4].text = format_euro((val_A_m/tot_ita)*100) if tot_ita else "0,00"
            row_cells[5].text = f"{val_B_m:,}".replace(',', '.')
            row_cells[6].text = format_euro((val_B_m/tot_ita)*100) if tot_ita else "0,00"
            row_cells[7].text = f"{val_C_m:,}".replace(',', '.')
            row_cells[8].text = format_euro((val_C_m/tot_ita)*100) if tot_ita else "0,00"
            
            for cell in row_cells: cell.paragraphs[0].runs[0].bold = True

        # Riga Totale Italia (Grassetto)
        row_cells = t6.add_row().cells
        row_cells[0].text = "ITALIA (TOTALE)"
        row_cells[1].text = f"{tot_ita:,}".replace(',', '.')
        row_cells[2].text = "100,00"
        row_cells[3].text = f"{tot_ita_A:,}".replace(',', '.')
        row_cells[4].text = format_euro((tot_ita_A/tot_ita)*100) if tot_ita else "0,00"
        row_cells[5].text = f"{tot_ita_B:,}".replace(',', '.')
        row_cells[6].text = format_euro((tot_ita_B/tot_ita)*100) if tot_ita else "0,00"
        row_cells[7].text = f"{tot_ita_C:,}".replace(',', '.')
        row_cells[8].text = format_euro((tot_ita_C/tot_ita)*100) if tot_ita else "0,00"
        
        for cell in row_cells: cell.paragraphs[0].runs[0].bold = True

    context['tabella_6_dinamica'] = sd_tab6

    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 7 E GRAFICI (COMPOSIZIONE VdP)
    # =================================================================
    # 1. RECUPERO DATI (Calcolo incidenza percentuale corretto e ANTI-CRASH)
    anni = ['2021', '2022', '2023', '2024']
    metriche = [
        ('Costo del venduto', 'Costo del venduto migl EUR'),
        ('Oneri di gestione', 'Oneri diversi di gestione migl EUR'),
        ('Proventi/Oneri fin.', 'Proventi/oneri finanziari migl EUR'),
        ('Imposte', 'Totale imposte migl EUR'),
        ('Utile/Perdita Netta', 'Utile/Perdita al netto delle imposte migl EUR')
    ]
    col_prod_prefisso = 'Totale valore della produzione migl EUR'

    dati_settore = {m[0]: [] for m in metriche}
    dati_azienda = {m[0]: [] for m in metriche}

    for nome_metrica, col_base in metriche:
        for anno in anni:
            col_num = f"{col_base} {anno}"
            col_den = f"{col_prod_prefisso} {anno}"

            # --- MEDIANA SETTORE (Blindata contro Zeri e Lettere) ---
            s_num = pd.to_numeric(df_orbis[col_num] if col_num in df_orbis.columns else pd.Series(dtype=float), errors='coerce')
            s_den = pd.to_numeric(df_orbis[col_den] if col_den in df_orbis.columns else pd.Series(dtype=float), errors='coerce')

            if not s_num.empty and not s_den.empty:
                # Trasforma gli zeri in NaN per non far esplodere la divisione
                s_den_safe = s_den.replace(0, np.nan)
                pct_series = (s_num / s_den_safe).replace([np.inf, -np.inf], np.nan).dropna()
                val_sett_pct = (pct_series.median() * 100) if not pct_series.empty else 0.0
            else:
                val_sett_pct = 0.0
            
            # Filtro di sicurezza estrema prima del grafico
            if pd.isna(val_sett_pct) or val_sett_pct == np.inf or val_sett_pct == -np.inf:
                val_sett_pct = 0.0
                
            dati_settore[nome_metrica].append(val_sett_pct)

            # --- VALORE AZIENDA TARGET (Blindato contro Zeri e Lettere) ---
            if not df_target.empty and col_num in df_target.columns and col_den in df_target.columns:
                try:
                    v_num = float(df_target.iloc[0].get(col_num, 0))
                    v_den = float(df_target.iloc[0].get(col_den, 1))
                    
                    if pd.isna(v_num): v_num = 0.0
                    if pd.isna(v_den) or v_den == 0: v_den = 1.0 # Evita div/0
                    
                    val_az_pct = (v_num / v_den) * 100
                except:
                    val_az_pct = 0.0
            else:
                val_az_pct = 0.0

            # Filtro di sicurezza estrema prima del grafico
            if pd.isna(val_az_pct) or val_az_pct == np.inf or val_az_pct == -np.inf:
                val_az_pct = 0.0

            dati_azienda[nome_metrica].append(val_az_pct)
            

    # 2. COSTRUZIONE TABELLA 7 (Word)
    sd_tab7 = doc.new_subdoc()
    t7 = sd_tab7.add_table(rows=1, cols=5)
    t7.style = 'Table Grid'

    # Intestazioni
    headers_t7 = ['Metrica / Componente %', '2021', '2022', '2023', '2024']
    for i, h in enumerate(headers_t7):
        t7.cell(0, i).text = h
        t7.cell(0, i).paragraphs[0].runs[0].bold = True

    # Inserimento righe alternate (Settore / Azienda)
    for nome_metrica, _ in metriche:
        # Riga Settore
        row_sett = t7.add_row().cells
        row_sett[0].text = f"{nome_metrica} (Mediane Settore)"
        for i, val in enumerate(dati_settore[nome_metrica]):
            row_sett[i+1].text = f"{val:.2f}%".replace('.', ',') if pd.notna(val) else "N.D."
            
        # Riga Azienda
        row_az = t7.add_row().cells
        row_az[0].text = f"{nome_metrica} ({context.get('ragione_sociale', 'Azienda')})"
        row_az[0].paragraphs[0].runs[0].bold = True # Evidenziamo il nome azienda
        for i, val in enumerate(dati_azienda[nome_metrica]):
            row_az[i+1].text = f"{val:.2f}%".replace('.', ',') if pd.notna(val) else "N.D."

    context['tabella_7_dinamica'] = sd_tab7

    # 3. GENERAZIONE GRAFICI (Matplotlib)
    def genera_grafico_barre_impilate(dati, titolo):
        fig, ax = plt.subplots(figsize=(8, 4.5))
        
        # Colori professionali (toni di blu, grigio, verde, ecc.)
        colori = ['#2b3a67', '#496a81', '#66999b', '#b3af8f', '#ffc482']
        
        bottoms = np.zeros(len(anni))
        for i, (nome, valori) in enumerate(dati.items()):
            # Sostituiamo eventuali NaN con 0 per non far crashare il grafico
            valori_puliti = [v if pd.notna(v) else 0 for v in valori]
            ax.bar(anni, valori_puliti, bottom=bottoms, label=nome, color=colori[i], edgecolor='white')
            bottoms += np.array(valori_puliti)
            
        ax.set_title(titolo, fontsize=12, fontweight='bold', color='#333333')
        ax.set_ylabel('% sul Valore della Produzione', fontsize=10)
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1), frameon=False)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(axis='y', linestyle='--', alpha=0.7)
        #plt.tight_layout()
        
        # Salviamo l'immagine in memoria (senza creare file sul pc)
        mem_img = io.BytesIO()

        aggiungi_watermark_fig(fig, modalita_teaser) # <--- AGGIUNGI ", modalita_teaser"

        plt.savefig(mem_img, format='png', dpi=300, bbox_inches='tight')
        plt.close(fig)
        mem_img.seek(0)
        return mem_img

    # Creazione Immagini e Iniezione nel Word
    img_settore_buf = genera_grafico_barre_impilate(dati_settore, "Composizione % - Mediane di Settore")
    context['grafico_settore'] = InlineImage(doc, img_settore_buf, width=Mm(155))

    img_azienda_buf = genera_grafico_barre_impilate(dati_azienda, f"Composizione % - {context.get('ragione_sociale', 'Azienda')}")
    context['grafico_azienda'] = InlineImage(doc, img_azienda_buf, width=Mm(155))
    # =================================================================

    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 8 E GRAFICO 3 (TREND MARGINE EBITDA)
    # =================================================================
    
    anni = ['2021', '2022', '2023', '2024']
    col_base_ebitda = 'Margine EBITDA (*) %'
    
    valori_settore_ebitda = []
    valori_azienda_ebitda = []
    
    # 1. Recupero Dati
    for anno in anni:
        colonna = f"{col_base_ebitda} {anno}"
        
        # Mediana Settore
        if colonna in df_orbis.columns:
            valori_settore_ebitda.append(df_orbis[colonna].median())
        else:
            valori_settore_ebitda.append(0.0)
            
        # Valore Azienda
        if not df_target.empty and colonna in df_target.columns:
            valori_azienda_ebitda.append(df_target.iloc[0][colonna])
        else:
            valori_azienda_ebitda.append(0.0)

    # 2. Costruzione Tabella 8
    sd_tab8 = doc.new_subdoc()
    t8 = sd_tab8.add_table(rows=3, cols=5)
    t8.style = 'Table Grid'
    
    # Intestazioni anni
    headers_t8 = ['', '2021', '2022', '2023', '2024']
    for i, h in enumerate(headers_t8):
        t8.cell(0, i).text = h
        t8.cell(0, i).paragraphs[0].runs[0].bold = True
        
    nome_azienda = context.get('ragione_sociale', 'Azienda Target')
    codice_nace_report = context.get('codice_nace', 'N.D.')
    
    # Riga Azienda
    t8.cell(1, 0).text = nome_azienda
    t8.cell(1, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_azienda_ebitda):
        t8.cell(1, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    # Riga Settore (Mediana)
    t8.cell(2, 0).text = f"Settore {codice_nace_report} (Mediana)"
    t8.cell(2, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_settore_ebitda):
        t8.cell(2, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    context['tabella_8_dinamica'] = sd_tab8

    # 3. Funzione per generare Grafici di Trend (Linee)
    def genera_grafico_trend(anni, val_azienda, val_settore, nome_az, titolo):
        fig, ax = plt.subplots(figsize=(8, 4.5))
        
        # Pulizia dati per il grafico (sostituiamo eventuali NaN con 0)
        val_az_clean = [v if pd.notna(v) else 0 for v in val_azienda]
        val_set_clean = [v if pd.notna(v) else 0 for v in val_settore]
        
        # Disegno delle due linee
        ax.plot(anni, val_az_clean, marker='o', linewidth=2.5, markersize=8, color='#2b3a67', label=nome_az)
        ax.plot(anni, val_set_clean, marker='s', linewidth=2, markersize=7, color='#b3af8f', linestyle='--', label='Mediana Settore')
        
        ax.set_title(titolo, fontsize=12, fontweight='bold', color='#333333')
        ax.set_ylabel('%', fontsize=10)
        ax.legend(loc='best', frameon=False)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(axis='y', linestyle='--', alpha=0.5)
        #plt.tight_layout()
        
        # Salvataggio in memoria
        mem_img = io.BytesIO()

        aggiungi_watermark_fig(fig, modalita_teaser) # <--- AGGIUNGI ", modalita_teaser"

        plt.savefig(mem_img, format='png', dpi=300, bbox_inches='tight')
        plt.close(fig)
        mem_img.seek(0)
        return mem_img

    # Generazione Immagine e Iniezione nel Word
    img_ebitda_buf = genera_grafico_trend(
        anni, valori_azienda_ebitda, valori_settore_ebitda, 
        nome_azienda, "Andamento Margine EBITDA (%)"
    )
    context['grafico_ebitda'] = InlineImage(doc, img_ebitda_buf, width=Mm(155))

    # =================================================================
    # 🏗️ COSTRUZIONE GRAFICO 4 (CONFRONTO DIRETTO EBITDA 2024)
    # =================================================================
    
    # Funzione riutilizzabile per grafici a barre comparative (Azienda vs Settore per 1 singolo anno)
    def genera_grafico_confronto_singolo(val_az, val_set, nome_az, titolo):
        fig, ax = plt.subplots(figsize=(8, 5.5))
        
        etichette = [nome_az, 'Mediana Settore']
        
        # Pulizia da eventuali NaN
        v_az_clean = val_az if pd.notna(val_az) else 0
        v_set_clean = val_set if pd.notna(val_set) else 0
        valori = [v_az_clean, v_set_clean]
        
        colori = ['#2b3a67', '#b3af8f']
        
        # Disegno delle barre
        bars = ax.bar(etichette, valori, color=colori, edgecolor='white', width=0.5)

        for bar in bars:
            yval = bar.get_height()
            if yval != 0:
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    yval / 2,
                    f"{yval:.2f}%".replace('.', ','),
                    ha='center', va='center',
                    fontsize=9, fontweight='bold', color='white'
                )
        
            
        ax.set_title(titolo, fontsize=11, fontweight='bold', color='#333333')
        ax.set_ylabel('%', fontsize=10)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(axis='y', linestyle='--', alpha=0.5)
        #plt.tight_layout()
        
        # Salvataggio in memoria
        mem_img = io.BytesIO()

        aggiungi_watermark_fig(fig, modalita_teaser) # <--- AGGIUNGI ", modalita_teaser"
        
        plt.savefig(mem_img, format='png', dpi=300, bbox_inches='tight')
        plt.close(fig)
        mem_img.seek(0)
        return mem_img

    # Prendiamo i dati del 2024 (che è l'ultimo valore della lista calcolata prima, quindi indice -1)
    val_az_ebitda_24 = valori_azienda_ebitda[-1] if len(valori_azienda_ebitda) > 0 else 0
    val_set_ebitda_24 = valori_settore_ebitda[-1] if len(valori_settore_ebitda) > 0 else 0

    # Generazione Immagine e Iniezione nel Word
    img_ebitda_24_buf = genera_grafico_confronto_singolo(
        val_az_ebitda_24, val_set_ebitda_24, nome_azienda, "Margine EBITDA (%) - 2024"
    )
    context['grafico_ebitda_2024'] = InlineImage(doc, img_ebitda_24_buf, width=Mm(125))


    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 9 E GRAFICO 5 (TREND MARGINE EBIT)
    # =================================================================
    
    col_base_ebit = 'Margine EBIT (*) %'
    
    valori_settore_ebit = []
    valori_azienda_ebit = []
    
    # 1. Recupero Dati
    for anno in anni: # 'anni' è la lista ['2021', '2022', '2023', '2024'] già definita sopra
        colonna = f"{col_base_ebit} {anno}"
        
        # Mediana Settore
        if colonna in df_orbis.columns:
            valori_settore_ebit.append(df_orbis[colonna].median())
        else:
            valori_settore_ebit.append(0.0)
            
        # Valore Azienda
        if not df_target.empty and colonna in df_target.columns:
            valori_azienda_ebit.append(df_target.iloc[0][colonna])
        else:
            valori_azienda_ebit.append(0.0)

    # 2. Costruzione Tabella 9
    sd_tab9 = doc.new_subdoc()
    t9 = sd_tab9.add_table(rows=3, cols=5)
    t9.style = 'Table Grid'
    
    # Intestazioni anni (riutilizziamo l'array headers_t8 che avevamo già)
    for i, h in enumerate(headers_t8):
        t9.cell(0, i).text = h
        t9.cell(0, i).paragraphs[0].runs[0].bold = True
        
    # Riga Azienda
    t9.cell(1, 0).text = nome_azienda
    t9.cell(1, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_azienda_ebit):
        t9.cell(1, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    # Riga Settore (Mediana)
    t9.cell(2, 0).text = f"Settore {codice_nace_report} (Mediana)"
    t9.cell(2, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_settore_ebit):
        t9.cell(2, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    context['tabella_9_dinamica'] = sd_tab9

    # 3. Generazione Immagine e Iniezione nel Word (Riutilizziamo la funzione magica!)
    img_ebit_buf = genera_grafico_trend(
        anni, valori_azienda_ebit, valori_settore_ebit, 
        nome_azienda, "Andamento Margine EBIT (%)"
    )
    context['grafico_ebit'] = InlineImage(doc, img_ebit_buf, width=Mm(155))


    # =================================================================
    # 🏗️ COSTRUZIONE GRAFICO 6 (CONFRONTO DIRETTO EBIT 2024)
    # =================================================================
    
    # Prendiamo i dati del 2024 (l'ultimo valore della lista calcolata nel blocco precedente)
    val_az_ebit_24 = valori_azienda_ebit[-1] if len(valori_azienda_ebit) > 0 else 0
    val_set_ebit_24 = valori_settore_ebit[-1] if len(valori_settore_ebit) > 0 else 0

    # Generazione Immagine e Iniezione nel Word (riutilizziamo la funzione a barre)
    img_ebit_24_buf = genera_grafico_confronto_singolo(
        val_az_ebit_24, val_set_ebit_24, nome_azienda, "Margine EBIT (%) - 2024"
    )
    context['grafico_ebit_2024'] = InlineImage(doc, img_ebit_24_buf, width=Mm(125))


    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 10 E GRAFICO 7 (TREND MARGINE DI PROFITTO)
    # =================================================================
    
    col_base_profitto = 'Margine di Profitto (*) %' # Assicurati che nel tuo Excel si chiami così
    
    valori_settore_profitto = []
    valori_azienda_profitto = []
    
    # 1. Recupero Dati
    for anno in anni:
        colonna = f"{col_base_profitto} {anno}"
        
        # Mediana Settore
        if colonna in df_orbis.columns:
            valori_settore_profitto.append(df_orbis[colonna].median())
        else:
            valori_settore_profitto.append(0.0)
            
        # Valore Azienda
        if not df_target.empty and colonna in df_target.columns:
            valori_azienda_profitto.append(df_target.iloc[0][colonna])
        else:
            valori_azienda_profitto.append(0.0)

    # 2. Costruzione Tabella 10
    sd_tab10 = doc.new_subdoc()
    t10 = sd_tab10.add_table(rows=3, cols=5)
    t10.style = 'Table Grid'
    
    # Intestazioni anni (questa volta la prima cella ha del testo)
    headers_t10 = ['Margine di Profitto', '2021', '2022', '2023', '2024']
    for i, h in enumerate(headers_t10):
        t10.cell(0, i).text = h
        t10.cell(0, i).paragraphs[0].runs[0].bold = True
        
    # Riga Azienda
    t10.cell(1, 0).text = nome_azienda
    t10.cell(1, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_azienda_profitto):
        t10.cell(1, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    # Riga Settore (Mediana)
    t10.cell(2, 0).text = f"Settore {codice_nace_report} (Mediana)"
    t10.cell(2, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_settore_profitto):
        t10.cell(2, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    context['tabella_10_dinamica'] = sd_tab10

    # 3. Generazione Immagine e Iniezione nel Word
    img_profitto_buf = genera_grafico_trend(
        anni, valori_azienda_profitto, valori_settore_profitto, 
        nome_azienda, "Andamento Margine di Profitto (%)"
    )
    context['grafico_profitto'] = InlineImage(doc, img_profitto_buf, width=Mm(155))

    # =================================================================
    # 🏗️ COSTRUZIONE GRAFICO 8 (CONFRONTO DIRETTO PROFITTO 2024)
    # =================================================================
    
    # Prendiamo i dati del 2024 (l'ultimo valore della lista calcolata nel blocco precedente)
    val_az_profitto_24 = valori_azienda_profitto[-1] if len(valori_azienda_profitto) > 0 else 0
    val_set_profitto_24 = valori_settore_profitto[-1] if len(valori_settore_profitto) > 0 else 0

    # Generazione Immagine e Iniezione nel Word (riutilizziamo la funzione a barre)
    img_profitto_24_buf = genera_grafico_confronto_singolo(
        val_az_profitto_24, val_set_profitto_24, nome_azienda, "Margine di Profitto (%) - 2024"
    )
    context['grafico_profitto_2024'] = InlineImage(doc, img_profitto_24_buf, width=Mm(125))

    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 11 E GRAFICO 9 (TREND STRUTTURA 1° LIV)
    # =================================================================
    
    col_base_strut1 = 'Indice di Struttura 1° livello (*)' # Verifica l'esatto nome nel tuo Excel senza l'anno
    
    valori_settore_strut1 = []
    valori_azienda_strut1 = []
    
    # 1. Recupero Dati
    for anno in anni:
        colonna = f"{col_base_strut1} {anno}"
        
        # Mediana Settore
        if colonna in df_orbis.columns:
            valori_settore_strut1.append(df_orbis[colonna].median())
        else:
            valori_settore_strut1.append(0.0)
            
        # Valore Azienda
        if not df_target.empty and colonna in df_target.columns:
            valori_azienda_strut1.append(df_target.iloc[0][colonna])
        else:
            valori_azienda_strut1.append(0.0)

    # 2. Costruzione Tabella 11
    sd_tab11 = doc.new_subdoc()
    t11 = sd_tab11.add_table(rows=3, cols=5)
    t11.style = 'Table Grid'
    
    # Intestazioni anni 
    headers_t11 = ['Indice Struttura 1° Liv.', '2021', '2022', '2023', '2024']
    for i, h in enumerate(headers_t11):
        t11.cell(0, i).text = h
        t11.cell(0, i).paragraphs[0].runs[0].bold = True
        
    # Riga Azienda
    t11.cell(1, 0).text = nome_azienda
    t11.cell(1, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_azienda_strut1):
        t11.cell(1, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    # Riga Settore (Mediana)
    t11.cell(2, 0).text = f"Settore {codice_nace_report} (Mediana)"
    t11.cell(2, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_settore_strut1):
        t11.cell(2, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    context['tabella_11_dinamica'] = sd_tab11

    # 3. Generazione Immagine e Iniezione nel Word
    img_strut1_buf = genera_grafico_trend(
        anni, valori_azienda_strut1, valori_settore_strut1, 
        nome_azienda, "Andamento Indice di Struttura 1° Liv."
    )
    context['grafico_strut1'] = InlineImage(doc, img_strut1_buf, width=Mm(155))


    # =================================================================
    # 🏗️ COSTRUZIONE GRAFICO 10 (CONFRONTO DIRETTO STRUTTURA 1° LIV 2024)
    # =================================================================
    
    # Prendiamo i dati del 2024 (l'ultimo valore della lista calcolata nel blocco precedente)
    val_az_strut1_24 = valori_azienda_strut1[-1] if len(valori_azienda_strut1) > 0 else 0
    val_set_strut1_24 = valori_settore_strut1[-1] if len(valori_settore_strut1) > 0 else 0

    # Generazione Immagine e Iniezione nel Word (riutilizziamo la funzione a barre)
    img_strut1_24_buf = genera_grafico_confronto_singolo(
        val_az_strut1_24, val_set_strut1_24, nome_azienda, "Indice Struttura 1° Liv. - 2024"
    )
    context['grafico_strut1_2024'] = InlineImage(doc, img_strut1_24_buf, width=Mm(125))

    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 12 E GRAFICO 11 (TREND STRUTTURA 2° LIV)
    # =================================================================
    
    col_base_strut2 = 'Indice di Struttura 2° livello (*)' # Verifica l'esatto nome nel tuo Excel senza l'anno
    
    valori_settore_strut2 = []
    valori_azienda_strut2 = []
    
    # 1. Recupero Dati
    for anno in anni:
        colonna = f"{col_base_strut2} {anno}"
        
        # Mediana Settore
        if colonna in df_orbis.columns:
            valori_settore_strut2.append(df_orbis[colonna].median())
        else:
            valori_settore_strut2.append(0.0)
            
        # Valore Azienda
        if not df_target.empty and colonna in df_target.columns:
            valori_azienda_strut2.append(df_target.iloc[0][colonna])
        else:
            valori_azienda_strut2.append(0.0)

    # 2. Costruzione Tabella 12
    sd_tab12 = doc.new_subdoc()
    t12 = sd_tab12.add_table(rows=3, cols=5)
    t12.style = 'Table Grid'
    
    # Intestazioni anni 
    headers_t12 = ['Indice Struttura 2° Liv.', '2021', '2022', '2023', '2024']
    for i, h in enumerate(headers_t12):
        t12.cell(0, i).text = h
        t12.cell(0, i).paragraphs[0].runs[0].bold = True
        
    # Riga Azienda
    t12.cell(1, 0).text = nome_azienda
    t12.cell(1, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_azienda_strut2):
        t12.cell(1, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    # Riga Settore (Mediana)
    t12.cell(2, 0).text = f"Settore {codice_nace_report} (Mediana)"
    t12.cell(2, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_settore_strut2):
        t12.cell(2, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    context['tabella_12_dinamica'] = sd_tab12

    # 3. Generazione Immagine e Iniezione nel Word
    img_strut2_buf = genera_grafico_trend(
        anni, valori_azienda_strut2, valori_settore_strut2, 
        nome_azienda, "Andamento Indice di Struttura 2° Liv."
    )
    context['grafico_strut2'] = InlineImage(doc, img_strut2_buf, width=Mm(155))

    # =================================================================
    # 🏗️ COSTRUZIONE GRAFICO 12 (CONFRONTO DIRETTO STRUTTURA 2° LIV 2024)
    # =================================================================
    
    # Prendiamo i dati del 2024 (l'ultimo valore della lista calcolata nel blocco precedente)
    val_az_strut2_24 = valori_azienda_strut2[-1] if len(valori_azienda_strut2) > 0 else 0
    val_set_strut2_24 = valori_settore_strut2[-1] if len(valori_settore_strut2) > 0 else 0

    # Generazione Immagine e Iniezione nel Word (riutilizziamo la funzione a barre)
    img_strut2_24_buf = genera_grafico_confronto_singolo(
        val_az_strut2_24, val_set_strut2_24, nome_azienda, "Indice Struttura 2° Liv. - 2024"
    )
    context['grafico_strut2_2024'] = InlineImage(doc, img_strut2_24_buf, width=Mm(125))

    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 13 E GRAFICO 13 (TREND GEARING)
    # =================================================================
    
    col_base_gearing = 'Gearing (*) %' # Verifica l'esatto nome nel tuo Excel senza l'anno
    
    valori_settore_gearing = []
    valori_azienda_gearing = []
    
    # 1. Recupero Dati
    for anno in anni:
        colonna = f"{col_base_gearing} {anno}"
        
        # Mediana Settore
        if colonna in df_orbis.columns:
            valori_settore_gearing.append(df_orbis[colonna].median())
        else:
            valori_settore_gearing.append(0.0)
            
        # Valore Azienda
        if not df_target.empty and colonna in df_target.columns:
            valori_azienda_gearing.append(df_target.iloc[0][colonna])
        else:
            valori_azienda_gearing.append(0.0)

    # 2. Costruzione Tabella 13
    sd_tab13 = doc.new_subdoc()
    t13 = sd_tab13.add_table(rows=3, cols=5)
    t13.style = 'Table Grid'
    
    # Intestazioni anni 
    headers_t13 = ['Gearing', '2021', '2022', '2023', '2024']
    for i, h in enumerate(headers_t13):
        t13.cell(0, i).text = h
        t13.cell(0, i).paragraphs[0].runs[0].bold = True
        
    # Riga Azienda
    t13.cell(1, 0).text = nome_azienda
    t13.cell(1, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_azienda_gearing):
        t13.cell(1, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    # Riga Settore (Mediana)
    t13.cell(2, 0).text = f"Settore {codice_nace_report} (Mediana)"
    t13.cell(2, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_settore_gearing):
        t13.cell(2, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    context['tabella_13_dinamica'] = sd_tab13

    # 3. Generazione Immagine e Iniezione nel Word
    img_gearing_buf = genera_grafico_trend(
        anni, valori_azienda_gearing, valori_settore_gearing, 
        nome_azienda, "Andamento Gearing (%)"
    )
    context['grafico_gearing'] = InlineImage(doc, img_gearing_buf, width=Mm(155))

    # =================================================================
    # 🏗️ COSTRUZIONE GRAFICO 14 (CONFRONTO DIRETTO GEARING 2024)
    # =================================================================
    
    # Prendiamo i dati del 2024 (l'ultimo valore della lista calcolata nel blocco precedente)
    val_az_gearing_24 = valori_azienda_gearing[-1] if len(valori_azienda_gearing) > 0 else 0
    val_set_gearing_24 = valori_settore_gearing[-1] if len(valori_settore_gearing) > 0 else 0

    # Generazione Immagine e Iniezione nel Word (riutilizziamo la funzione a barre)
    img_gearing_24_buf = genera_grafico_confronto_singolo(
        val_az_gearing_24, val_set_gearing_24, nome_azienda, "Gearing (%) - 2024"
    )
    context['grafico_gearing_2024'] = InlineImage(doc, img_gearing_24_buf, width=Mm(125))

    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 14 E GRAFICO TREND (CURRENT RATIO)
    # =================================================================
    
    col_base_cr = 'Current Ratio (*)' # Controlla che nel tuo Excel si chiami così o aggiungi la "x" se c'è
    
    valori_settore_cr = []
    valori_azienda_cr = []
    
    # 1. Recupero Dati
    for anno in anni:
        colonna = f"{col_base_cr} {anno}"
        
        # Mediana Settore
        if colonna in df_orbis.columns:
            valori_settore_cr.append(df_orbis[colonna].median())
        else:
            valori_settore_cr.append(0.0)
            
        # Valore Azienda
        if not df_target.empty and colonna in df_target.columns:
            valori_azienda_cr.append(df_target.iloc[0][colonna])
        else:
            valori_azienda_cr.append(0.0)

    # 2. Costruzione Tabella 14
    sd_tab14 = doc.new_subdoc()
    t14 = sd_tab14.add_table(rows=3, cols=5)
    t14.style = 'Table Grid'
    
    # Intestazioni anni 
    headers_t14 = ['Current Ratio', '2021', '2022', '2023', '2024']
    for i, h in enumerate(headers_t14):
        t14.cell(0, i).text = h
        t14.cell(0, i).paragraphs[0].runs[0].bold = True
        
    # Riga Azienda
    t14.cell(1, 0).text = nome_azienda
    t14.cell(1, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_azienda_cr):
        t14.cell(1, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    # Riga Settore (Mediana)
    t14.cell(2, 0).text = f"Settore {codice_nace_report} (Mediana)"
    t14.cell(2, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_settore_cr):
        t14.cell(2, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    context['tabella_14_dinamica'] = sd_tab14

    # 3. Generazione Immagine e Iniezione nel Word
    img_cr_buf = genera_grafico_trend(
        anni, valori_azienda_cr, valori_settore_cr, 
        nome_azienda, "Andamento Current Ratio"
    )
    context['grafico_cr'] = InlineImage(doc, img_cr_buf, width=Mm(155))

    # =================================================================
    # 🏗️ COSTRUZIONE GRAFICO 15 (CONFRONTO DIRETTO CURRENT RATIO 2024)
    # =================================================================
    
    # Prendiamo i dati del 2024 (l'ultimo valore della lista calcolata nel blocco precedente)
    val_az_cr_24 = valori_azienda_cr[-1] if len(valori_azienda_cr) > 0 else 0
    val_set_cr_24 = valori_settore_cr[-1] if len(valori_settore_cr) > 0 else 0

    # Generazione Immagine e Iniezione nel Word (riutilizziamo la funzione a barre)
    img_cr_24_buf = genera_grafico_confronto_singolo(
        val_az_cr_24, val_set_cr_24, nome_azienda, "Current Ratio - 2024"
    )
    context['grafico_cr_2024'] = InlineImage(doc, img_cr_24_buf, width=Mm(125))


    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 15 E GRAFICO 16 (TREND QUICK RATIO)
    # =================================================================
    
    col_base_qr = 'Quick Ratio (*)' # Controlla che nel tuo Excel la colonna si chiami così
    
    valori_settore_qr = []
    valori_azienda_qr = []
    
    # 1. Recupero Dati
    for anno in anni:
        colonna = f"{col_base_qr} {anno}"
        
        # Mediana Settore
        if colonna in df_orbis.columns:
            valori_settore_qr.append(df_orbis[colonna].median())
        else:
            valori_settore_qr.append(0.0)
            
        # Valore Azienda
        if not df_target.empty and colonna in df_target.columns:
            valori_azienda_qr.append(df_target.iloc[0][colonna])
        else:
            valori_azienda_qr.append(0.0)

    # 2. Costruzione Tabella 15
    sd_tab15 = doc.new_subdoc()
    t15 = sd_tab15.add_table(rows=3, cols=5)
    t15.style = 'Table Grid'
    
    # Intestazioni anni 
    headers_t15 = ['Quick Ratio', '2021', '2022', '2023', '2024']
    for i, h in enumerate(headers_t15):
        t15.cell(0, i).text = h
        t15.cell(0, i).paragraphs[0].runs[0].bold = True
        
    # Riga Azienda
    t15.cell(1, 0).text = nome_azienda
    t15.cell(1, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_azienda_qr):
        t15.cell(1, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    # Riga Settore (Mediana)
    t15.cell(2, 0).text = f"Settore {codice_nace_report} (Mediana)"
    t15.cell(2, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_settore_qr):
        t15.cell(2, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    context['tabella_15_dinamica'] = sd_tab15

    # 3. Generazione Immagine e Iniezione nel Word
    img_qr_buf = genera_grafico_trend(
        anni, valori_azienda_qr, valori_settore_qr, 
        nome_azienda, "Andamento Quick Ratio"
    )
    context['grafico_qr'] = InlineImage(doc, img_qr_buf, width=Mm(155))


    # =================================================================
    # 🏗️ COSTRUZIONE GRAFICO 17 (CONFRONTO DIRETTO QUICK RATIO 2024)
    # =================================================================
    
    # Prendiamo i dati del 2024 (l'ultimo valore della lista calcolata nel blocco precedente)
    val_az_qr_24 = valori_azienda_qr[-1] if len(valori_azienda_qr) > 0 else 0
    val_set_qr_24 = valori_settore_qr[-1] if len(valori_settore_qr) > 0 else 0

    # Generazione Immagine e Iniezione nel Word (riutilizziamo la funzione a barre)
    img_qr_24_buf = genera_grafico_confronto_singolo(
        val_az_qr_24, val_set_qr_24, nome_azienda, "Quick Ratio - 2024"
    )
    context['grafico_qr_2024'] = InlineImage(doc, img_qr_24_buf, width=Mm(125))


    # =================================================================
    # 🏗️ COSTRUZIONE TABELLA 16 E GRAFICO 18 (TREND ROTAZIONE CAP. INV.)
    # =================================================================
    
    # IMPORTANTE: Controlla che il nome di questa colonna corrisponda esattamente a quello del tuo Excel Orbis!
    col_base_rotazione = 'Indice di Rotazione del Capitale Investito (*)' 
    
    valori_settore_rotazione = []
    valori_azienda_rotazione = []
    
    # 1. Recupero Dati
    for anno in anni:
        colonna = f"{col_base_rotazione} {anno}"
        
        # Mediana Settore
        if colonna in df_orbis.columns:
            valori_settore_rotazione.append(df_orbis[colonna].median())
        else:
            valori_settore_rotazione.append(0.0)
            
        # Valore Azienda
        if not df_target.empty and colonna in df_target.columns:
            valori_azienda_rotazione.append(df_target.iloc[0][colonna])
        else:
            valori_azienda_rotazione.append(0.0)

    # 2. Costruzione Tabella 16
    sd_tab16 = doc.new_subdoc()
    t16 = sd_tab16.add_table(rows=3, cols=5)
    t16.style = 'Table Grid'
    
    # Intestazioni anni 
    headers_t16 = ['Indice Rotazione Cap.Inv.', '2021', '2022', '2023', '2024']
    for i, h in enumerate(headers_t16):
        t16.cell(0, i).text = h
        t16.cell(0, i).paragraphs[0].runs[0].bold = True
        
    # Riga Azienda
    t16.cell(1, 0).text = nome_azienda
    t16.cell(1, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_azienda_rotazione):
        t16.cell(1, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    # Riga Settore (Mediana)
    t16.cell(2, 0).text = f"Settore {codice_nace_report} (Mediana)"
    t16.cell(2, 0).paragraphs[0].runs[0].bold = True
    for i, val in enumerate(valori_settore_rotazione):
        t16.cell(2, i+1).text = f"{val:.2f}".replace('.', ',') if pd.notna(val) else "N.D."
        
    context['tabella_16_dinamica'] = sd_tab16

    # 3. Generazione Immagine e Iniezione nel Word
    img_rotazione_buf = genera_grafico_trend(
        anni, valori_azienda_rotazione, valori_settore_rotazione, 
        nome_azienda, "Andamento Rotazione Cap. Inv."
    )
    context['grafico_rotazione'] = InlineImage(doc, img_rotazione_buf, width=Mm(155))

    # =================================================================
    # 🏗️ COSTRUZIONE GRAFICO 19 (CONFRONTO DIRETTO ROTAZIONE CAP. INV. 2024)
    # =================================================================
    
    # Prendiamo i dati del 2024 (l'ultimo valore della lista calcolata nel blocco precedente)
    val_az_rotazione_24 = valori_azienda_rotazione[-1] if len(valori_azienda_rotazione) > 0 else 0
    val_set_rotazione_24 = valori_settore_rotazione[-1] if len(valori_settore_rotazione) > 0 else 0

    # Generazione Immagine e Iniezione nel Word (riutilizziamo la funzione a barre)
    img_rotazione_24_buf = genera_grafico_confronto_singolo(
        val_az_rotazione_24, val_set_rotazione_24, nome_azienda, "Rotazione Cap. Inv. - 2024"
    )
    context['grafico_rotazione_2024'] = InlineImage(doc, img_rotazione_24_buf, width=Mm(125))

    # =================================================================
    # 🏗️ NOTA METODOLOGICA: CALCOLO E COSTRUZIONE TABELLE TERZILI 2024
    # =================================================================

    # 1. Funzione intelligente per calcolare e formattare i terzili di una colonna
    def calcola_terzili(df, col_name):
        if col_name not in df.columns:
            return ["N.D.", "N.D.", "N.D."]
        
        # Prendiamo solo i dati validi (senza NaN)
        s = df[col_name].dropna()
        if s.empty:
            return ["N.D.", "N.D.", "N.D."]
            
        vmin = s.min()
        v33 = s.quantile(0.3333)
        v66 = s.quantile(0.6666)
        vmax = s.max()
        
        # Funzione per formattare i numeri in stile italiano (1.234,56)
        def fmt(val):
            return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

        # Creazione delle tre stringhe per i terzili
        t1 = f"[{fmt(vmin)} - {fmt(v33)})"
        t2 = f"[{fmt(v33)} - {fmt(v66)})"
        t3 = f"[{fmt(v66)} - {fmt(vmax)}]"
        
        return [t1, t2, t3]

    # Funzione per costruire dinamicamente le tabelle Word dei terzili
    def crea_tabella_terzili(doc, headers, col_names_orbis):
        sd = doc.new_subdoc()
        t = sd.add_table(rows=4, cols=4)
        t.style = 'Table Grid'
        
        # Intestazioni (Riga 0)
        t.cell(0, 0).text = "Terzili"
        t.cell(0, 0).paragraphs[0].runs[0].bold = True
        for i, h in enumerate(headers):
            t.cell(0, i+1).text = h
            t.cell(0, i+1).paragraphs[0].runs[0].bold = True
            
        # Nomi delle righe
        t.cell(1, 0).text = "1°"
        t.cell(2, 0).text = "2°"
        t.cell(3, 0).text = "3°"
        
        # Popolamento dati incrociati
        for i, col_name in enumerate(col_names_orbis):
            fasce = calcola_terzili(df_orbis, col_name)
            t.cell(1, i+1).text = fasce[0]
            t.cell(2, i+1).text = fasce[1]
            t.cell(3, i+1).text = fasce[2]
            
        return sd

    # --- Creazione Tabella 1 (Profitto, EBITDA, EBIT) ---
    headers_t1 = ['Margine di profitto', 'Margine EBITDA', 'Margine EBIT']
    colonne_t1 = [f"{col_base_profitto} 2024", f"{col_base_ebitda} 2024", f"{col_base_ebit} 2024"]
    context['tabella_terzili_1'] = crea_tabella_terzili(doc, headers_t1, colonne_t1)

    # --- Creazione Tabella 2 (Struttura 1°, Struttura 2°, Gearing) ---
    headers_t2 = ['Indice Struttura 1° Livello', 'Indice Struttura 2° Livello', 'Indice Gearing']
    colonne_t2 = [f"{col_base_strut1} 2024", f"{col_base_strut2} 2024", f"{col_base_gearing} 2024"]
    context['tabella_terzili_2'] = crea_tabella_terzili(doc, headers_t2, colonne_t2)

    # --- Creazione Tabella 3 (Rotazione, Quick, Current) ---
    headers_t3 = ['Indice Rotazione Cap.Inv.', 'Quick Ratio', 'Current Ratio']
    colonne_t3 = [f"{col_base_rotazione} 2024", f"{col_base_qr} 2024", f"{col_base_cr} 2024"]
    context['tabella_terzili_3'] = crea_tabella_terzili(doc, headers_t3, colonne_t3)

    # =================================================================
    # 🏗️ COSTRUZIONE TABELLE BENCHMARK (TOTALMENTE DINAMICHE DA APP.PY)
    # =================================================================

    def crea_tabella_benchmark(doc, titolo_colonna, dict_valori):
        sd = doc.new_subdoc()
        t = sd.add_table(rows=5, cols=3)
        t.style = 'Table Grid'
        
        totale_va = sum(dict_valori.values())
        
        t.cell(0, 0).text = titolo_colonna
        t.cell(0, 0).paragraphs[0].runs[0].bold = True
        t.cell(0, 1).text = "V.A."
        t.cell(0, 1).paragraphs[0].runs[0].bold = True
        t.cell(0, 2).text = "%"
        t.cell(0, 2).paragraphs[0].runs[0].bold = True
        
        for i, key in enumerate(['A', 'B', 'C']):
            va = dict_valori.get(key, 0)
            perc = (va / totale_va * 100) if totale_va > 0 else 0
            t.cell(i+1, 0).text = key
            t.cell(i+1, 1).text = f"{va:,}".replace(',', '.') 
            t.cell(i+1, 2).text = f"{perc:.2f}".replace('.', ',') 
            
        t.cell(4, 0).text = "Totale"
        t.cell(4, 0).paragraphs[0].runs[0].bold = True
        t.cell(4, 1).text = f"{totale_va:,}".replace(',', '.')
        t.cell(4, 1).paragraphs[0].runs[0].bold = True
        t.cell(4, 2).text = "100,00" if totale_va > 0 else "0,00"
        t.cell(4, 2).paragraphs[0].runs[0].bold = True
        
        return sd

    # ---------------------------------------------------------
    # 📊 MOTORE DI CALCOLO PANDAS (REPLICA FORMULE EXCEL APP.PY)
    # ---------------------------------------------------------
    df_b = df_orbis.copy()

    # Funzione per assegnare i punteggi 1, 2, 3 basata sui terzili
    def calcola_punti(serie_col, higher_is_better=True):
        s = pd.to_numeric(df_b[serie_col], errors='coerce')
        t1 = s.quantile(1/3)
        t2 = s.quantile(2/3)
        
        def assegna(x):
            if pd.isna(x): return 1 # Default se mancante
            if higher_is_better:
                return 3 if x >= t2 else (2 if x >= t1 else 1)
            else:
                return 3 if x <= t1 else (2 if x <= t2 else 1)
        return s.apply(assegna)

    # 1. Benchmark Economico (Più alto è meglio)
    pt_prof = calcola_punti(f"{col_base_profitto} 2024", True)
    pt_ebitda = calcola_punti(f"{col_base_ebitda} 2024", True)
    pt_ebit = calcola_punti(f"{col_base_ebit} 2024", True)
    df_b['Score_Eco'] = pt_prof + pt_ebitda + pt_ebit
    df_b['Bench_Eco'] = df_b['Score_Eco'].apply(lambda x: 'A' if x >= 8 else ('B' if x >= 5 else 'C'))

    # 2. Benchmark Finanziario (Rotazione: Più basso è meglio - Quick/Current: Più alto è meglio)
    pt_rot = calcola_punti(f"{col_base_rotazione} 2024", True) # in app.py cond_H <= T1
    pt_qr = calcola_punti(f"{col_base_qr} 2024", True)
    pt_cr = calcola_punti(f"{col_base_cr} 2024", True)
    df_b['Score_Fin'] = pt_rot + pt_qr + pt_cr
    df_b['Bench_Fin'] = df_b['Score_Fin'].apply(lambda x: 'A' if x >= 8 else ('B' if x >= 5 else 'C'))

    # 3. Benchmark Patrimoniale (Struttura: Più alto è meglio - Gearing: Più basso è meglio)
    pt_s1 = calcola_punti(f"{col_base_strut1} 2024", True)
    pt_s2 = calcola_punti(f"{col_base_strut2} 2024", True)
    pt_gear = calcola_punti(f"{col_base_gearing} 2024", False) # in app.py cond_N <= T1
    df_b['Score_Pat'] = pt_s1 + pt_s2 + pt_gear
    df_b['Bench_Pat'] = df_b['Score_Pat'].apply(lambda x: 'A' if x >= 8 else ('B' if x >= 5 else 'C'))

    # 4. Benchmark Totale
    def val_lettera(l):
        return 3 if l == 'A' else (2 if l == 'B' else 1)
    
    df_b['Score_Tot'] = df_b['Bench_Eco'].apply(val_lettera) + df_b['Bench_Fin'].apply(val_lettera) + df_b['Bench_Pat'].apply(val_lettera)
    df_b['Bench_Tot'] = df_b['Score_Tot'].apply(lambda x: 'A' if x >= 8 else ('B' if x >= 5 else 'C'))

    # Estrazione dei conteggi esatti
    conteggi_eco = df_b['Bench_Eco'].value_counts().to_dict()
    conteggi_fin = df_b['Bench_Fin'].value_counts().to_dict()
    conteggi_pat = df_b['Bench_Pat'].value_counts().to_dict()
    conteggi_tot = df_b['Bench_Tot'].value_counts().to_dict()

    # Iniezione nel Word
    context['tab_bench_eco'] = crea_tabella_benchmark(doc, "Benchmark Economico", conteggi_eco)
    context['tab_bench_pat'] = crea_tabella_benchmark(doc, "Benchmark Patrimoniale", conteggi_pat)
    context['tab_bench_fin'] = crea_tabella_benchmark(doc, "Benchmark Finanziario", conteggi_fin)
    context['tab_bench_tot'] = crea_tabella_benchmark(doc, "Benchmark Totale", conteggi_tot)



    # =================================================================
    # 🏗️ COSTRUZIONE TABELLE STATISTICHE DESCRITTIVE (PROFITTO, EBITDA, EBIT)
    # =================================================================

    # Funzione universale per creare la tabella delle statistiche
    def crea_tabella_statistiche(doc, titolo, col_base, anni, df):
        sd = doc.new_subdoc()
        t = sd.add_table(rows=6, cols=len(anni) + 1)
        t.style = 'Table Grid'
        
        # Intestazione Colonne
        t.cell(0, 0).text = titolo
        t.cell(0, 0).paragraphs[0].runs[0].bold = True
        for i, anno in enumerate(anni):
            t.cell(0, i+1).text = str(anno)
            t.cell(0, i+1).paragraphs[0].runs[0].bold = True
            
        # Intestazione Righe
        labels = ["Media", "Mediana", "Asimmetria", "Curtosi", "Deviazione standard"]
        for r, label in enumerate(labels):
            t.cell(r+1, 0).text = label
            t.cell(r+1, 0).paragraphs[0].runs[0].bold = True
            
        # Funzione di formattazione italiana (1.234,56)
        def fmt(val):
            if pd.isna(val): return "N.D."
            return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

        # Calcolo dinamico per ogni anno
        for i, anno in enumerate(anni):
            col_name = f"{col_base} {anno}"
            if col_name in df.columns:
                # Prendiamo solo i numeri puliti (scartando i null/NaN)
                s = pd.to_numeric(df[col_name], errors='coerce').dropna()
                
                if not s.empty:
                    media = s.mean()
                    mediana = s.median()
                    asimmetria = s.skew()
                    curtosi = s.kurt()
                    dev_std = s.std()
                else:
                    media = mediana = asimmetria = curtosi = dev_std = float('nan')
            else:
                media = mediana = asimmetria = curtosi = dev_std = float('nan')
                
            # Scrittura nella griglia Word
            t.cell(1, i+1).text = fmt(media)
            t.cell(2, i+1).text = fmt(mediana)
            t.cell(3, i+1).text = fmt(asimmetria)
            t.cell(4, i+1).text = fmt(curtosi)
            t.cell(5, i+1).text = fmt(dev_std)
            
        return sd

    # ---------------------------------------------------------
    # Iniezione nel Word
    # Utilizziamo le variabili col_base_profitto, col_base_ebitda, col_base_ebit
    # definite in precedenza nei blocchi dei singoli indicatori
    # ---------------------------------------------------------
    
    context['tab_stat_profitto'] = crea_tabella_statistiche(doc, "Margine di Profitto", col_base_profitto, anni, df_orbis)
    context['tab_stat_ebitda'] = crea_tabella_statistiche(doc, "Margine EBITDA", col_base_ebitda, anni, df_orbis)
    context['tab_stat_ebit'] = crea_tabella_statistiche(doc, "Margine EBIT", col_base_ebit, anni, df_orbis)

    context['tab_stat_strut1'] = crea_tabella_statistiche(doc, "Indice Struttura 1° Liv.", col_base_strut1, anni, df_orbis)
    context['tab_stat_strut2'] = crea_tabella_statistiche(doc, "Indice Struttura 2° Liv.", col_base_strut2, anni, df_orbis)
    context['tab_stat_gearing'] = crea_tabella_statistiche(doc, "Gearing (%)", col_base_gearing, anni, df_orbis)

    context['tab_stat_cr'] = crea_tabella_statistiche(doc, "Current Ratio", col_base_cr, anni, df_orbis)
    context['tab_stat_qr'] = crea_tabella_statistiche(doc, "Quick Ratio", col_base_qr, anni, df_orbis)
    context['tab_stat_rotazione'] = crea_tabella_statistiche(doc, "Rotazione Cap. Inv.", col_base_rotazione, anni, df_orbis)

    
    # =================================================================
    # 🏗️ COSTRUZIONE TABELLE MEDIANE E VARIAZIONI YoY (Tab. 16 - 24)
    # =================================================================

    def crea_tabella_yoy(doc, titolo, col_base, anni, df):
        sd = doc.new_subdoc()
        # Calcoliamo il numero di colonne: Nome + Primo Anno + 2 colonne (V.m. e Delta) per ogni anno successivo
        num_cols = 2 + (len(anni) - 1) * 2 
        t = sd.add_table(rows=2, cols=num_cols)
        t.style = 'Table Grid'
        
        # --- Costruzione Intestazioni ---
        t.cell(0, 0).text = "Metrica"
        t.cell(0, 1).text = f"{anni[0]} V.m."
        col_idx = 2
        for anno in anni[1:]:
            t.cell(0, col_idx).text = f"{anno} V.m."
            t.cell(0, col_idx+1).text = "Δ %"
            col_idx += 2
            
        for i in range(num_cols):
            t.cell(0, i).paragraphs[0].runs[0].bold = True
            
        t.cell(1, 0).text = titolo
        t.cell(1, 0).paragraphs[0].runs[0].bold = True
        
        # Funzioni di formattazione
        def fmt_val(val):
            return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(val) else "N.D."
            
        prev_val = None
        
        # --- Inserimento Dati 1° Anno (es. 2021) ---
        col_name_0 = f"{col_base} {anni[0]}"
        if col_name_0 in df.columns:
            val_0 = df[col_name_0].median()
        else:
            val_0 = float('nan')
            
        t.cell(1, 1).text = fmt_val(val_0)
        prev_val = val_0
        
        # --- Inserimento Anni Successivi e Calcolo Delta % ---
        col_idx = 2
        for anno in anni[1:]:
            col_name = f"{col_base} {anno}"
            if col_name in df.columns:
                val = df[col_name].median()
            else:
                val = float('nan')
                
            t.cell(1, col_idx).text = fmt_val(val)
            
            # Calcolo formula Δ %
            if prev_val is not None and pd.notna(prev_val) and prev_val != 0 and pd.notna(val):
                delta = ((val - prev_val) / abs(prev_val)) * 100
                t.cell(1, col_idx+1).text = fmt_val(delta) # Usiamo la stessa formattazione
            else:
                t.cell(1, col_idx+1).text = "n.d."
                
            prev_val = val
            col_idx += 2
            
        return sd

    # ---------------------------------------------------------
    # Iniezione delle 9 tabelle finali nel Word
    # ---------------------------------------------------------
    
    # Equilibrio Economico
    context['tab_yoy_profitto'] = crea_tabella_yoy(doc, "Margine di Profitto", col_base_profitto, anni, df_orbis)
    context['tab_yoy_ebitda'] = crea_tabella_yoy(doc, "Margine EBITDA", col_base_ebitda, anni, df_orbis)
    context['tab_yoy_ebit'] = crea_tabella_yoy(doc, "Margine EBIT", col_base_ebit, anni, df_orbis)
    
    # Equilibrio Patrimoniale
    context['tab_yoy_strut1'] = crea_tabella_yoy(doc, "Indice Strut. 1° Liv.", col_base_strut1, anni, df_orbis)
    context['tab_yoy_strut2'] = crea_tabella_yoy(doc, "Indice Strut. 2° Liv.", col_base_strut2, anni, df_orbis)
    context['tab_yoy_gearing'] = crea_tabella_yoy(doc, "Gearing (%)", col_base_gearing, anni, df_orbis)
    
    # Equilibrio Finanziario
    context['tab_yoy_cr'] = crea_tabella_yoy(doc, "Current Ratio", col_base_cr, anni, df_orbis)
    context['tab_yoy_qr'] = crea_tabella_yoy(doc, "Quick Ratio", col_base_qr, anni, df_orbis)
    context['tab_yoy_rotazione'] = crea_tabella_yoy(doc, "Rotazione Cap. Inv.", col_base_rotazione, anni, df_orbis)

    #---

    # =================================================================
    # 🌟 VARIABILI SOTTO I GRAFICI (Ora calcolate nel punto giusto!)
    # =================================================================
    context['ind_gear'] = context.get('gearing', 'N.D.')
    context['med_mg_ebitda'] = format_euro(val_set_ebitda_24)
    context['med_mg_ebit'] = format_euro(val_set_ebit_24)
    context['med_mg_prof'] = format_euro(val_set_profitto_24)
    context['med_ind_str1'] = format_euro(val_set_strut1_24)
    context['med_ind_str2'] = format_euro(val_set_strut2_24)
    context['med_ind_gear'] = format_euro(val_set_gearing_24)
    context['med_ind_cr'] = format_euro(val_set_cr_24)
    context['med_ind_qr'] = format_euro(val_set_qr_24)
    context['med_ind_rot_cap'] = format_euro(val_set_rotazione_24)

    # --- INDENTAZIONE AUTOMATICA DEI TESTI DISCORSIVI ---
    chiavi_narrative = [
        'descr_rating_tot', 'descr_rating_eco', 'descr_rating_patr', 'descr_rating_fin', 'descr_sintesi',
        'intro_benchmark_eco', 'intro_benchmark_patr', 'intro_benchmark_fin', 'analisi_combinata', 
        'intro_divario_strutturale', 'sintesi_quadriennio_patr', 'sintesi_modello_operativo_fin',
        'analisi_margini_operativi', 'analisi_margine_profitto', 'analisi_indici_struttura', 'analisi_gearing',
        'analisi_rotazione', 'analisi_current_ratio', 'analisi_quick_ratio', 'analisi_posizionamento_fin',
        'sintesi_profilo_integrato', 'sintesi_posizionamento_lungo_periodo', 'conclusione_patrimoniale',
        'conclusione_economica', 'conclusione_finanziaria_dettaglio', 'raccomandazione_finale', 'impatto_territoriale'
    ]

    # =================================================================
    # 📏 FIX SPAZIATURA E INDENTAZIONE NARRATIVA
    # =================================================================
    from docx.shared import Cm

    # 1. Spazzoliamo tutti i paragrafi del documento (ignorando l'interno delle tabelle)
    for p in doc.docx.paragraphs:
        if p.text.strip(): 
            # Interlinea 1.5
            p.paragraph_format.line_spacing = 1.0
            
            # Controlliamo se è un paragrafo che contiene le nostre variabili narrative
            # (Se il paragrafo inizia con uno dei tag che abbiamo inserito)
            for k in chiavi_narrative:
                if "{{" + k + "}}" in p.text:
                    # Indenta l'intero blocco di 1 cm a sinistra
                    p.paragraph_format.left_indent = Cm(1.0)
                    break

    # =================================================================
    # 🚑 FIX TABELLE FINALE E CENSURA TEASER
    # =================================================================
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import RGBColor

    for chiave, elemento in context.items():
        if hasattr(elemento, 'add_paragraph') or type(elemento).__name__ == 'Subdoc':
            if hasattr(elemento, 'tables'):
                for t in elemento.tables:
                    # ---------------------------------------------------------
                    # 🔒 1. LOGICA DI CENSURA TABELLE
                    # ---------------------------------------------------------
                    if modalita_teaser and len(t.rows) > 0:
                        intestazioni = [cell.text.strip() for cell in t.rows[0].cells]
                        
                        colonne_da_censurare = []
                        for c_idx, testo in enumerate(intestazioni):
                            if '2024' in testo:
                                colonne_da_censurare.append(c_idx)
                                if c_idx + 1 < len(intestazioni) and 'Δ' in intestazioni[c_idx + 1]:
                                    colonne_da_censurare.append(c_idx + 1)
                        
                        if colonne_da_censurare:
                            for r_idx in range(1, len(t.rows)):
                                for c_idx in colonne_da_censurare:
                                    if c_idx < len(t.rows[r_idx].cells):
                                        t.rows[r_idx].cells[c_idx].text = " 🔒 PREMIUM "
                                        if len(t.rows[r_idx].cells[c_idx].paragraphs) > 0:
                                            run = t.rows[r_idx].cells[c_idx].paragraphs[0].runs[0]
                                            run.font.bold = True
                                            run.font.color.rgb = RGBColor(255, 0, 0)
                        
                        elif len(intestazioni) > 0 and intestazioni[0] == 'Terzili':
                            for riga in t.rows[1:]:
                                for c in range(1, len(riga.cells)):
                                    riga.cells[c].text = " 🔒 PREMIUM "
                                    if len(riga.cells[c].paragraphs) > 0:
                                        run = riga.cells[c].paragraphs[0].runs[0]
                                        run.font.bold = True
                                        run.font.color.rgb = RGBColor(255, 0, 0)

                        elif len(t.rows) > 2 and len(t.rows[2].cells) > 0 and 'Ranking' in t.rows[2].cells[0].text:
                            for r_idx in range(2, len(t.rows)):
                                for c_idx in range(1, len(t.rows[r_idx].cells)):
                                    t.rows[r_idx].cells[c_idx].text = " 🔒 PREMIUM "
                                    if len(t.rows[r_idx].cells[c_idx].paragraphs) > 0:
                                        run = t.rows[r_idx].cells[c_idx].paragraphs[0].runs[0]
                                        run.font.bold = True
                                        run.font.color.rgb = RGBColor(255, 0, 0)

                    # ---------------------------------------------------------
                    # 🎨 2. LOGICA DI LAYOUT E COLORI ORBIS
                    # ---------------------------------------------------------
                    try:
                        t.autofit = True
                        tbl_pr = t._tbl.tblPr
                        tbl_w = tbl_pr.xpath("./w:tblW")
                        if tbl_w:
                            tbl_w[0].set(qn('w:type'), 'auto')
                            tbl_w[0].set(qn('w:w'), '0')
                        t.alignment = WD_TABLE_ALIGNMENT.CENTER
                        
                        if len(t.rows) > 0:
                            for r_idx, row in enumerate(t.rows):
                                for cell in row.cells:
                                    for paragraph in cell.paragraphs:
                                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                                    if r_idx == 0:
                                        shd = parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="2B3A67"/>' % nsdecls('w'))
                                        cell._tc.get_or_add_tcPr().append(shd)
                                        for paragraph in cell.paragraphs:
                                            for run in paragraph.runs:
                                                run.font.color.rgb = RGBColor(255, 255, 255)
                    except:
                        pass

    # =================================================================
    # 🔒 3. CENSURA VARIABILI (Testo Puro = Zero Crash)
    # =================================================================
    if modalita_teaser:
        testo_censura_sicuro = " 🔒 PREMIUM "
        
        variabili_sensibili = [
            'mg_ebitda', 'mg_ebit', 'mg_prof',
            'ind_str1', 'ind_str2', 'gearing', 'ind_gear',
            'ind_rot_cap', 'ind_cr', 'ind_qr',
            'ricavi_mln', 'attivo_mln',
            'perc_ricavi_panel', 'perc_attivo_panel', 'perc_dip_area',
            'perc_imprese_regione', 'perc_ricavi_macroregione', 'tot_ricavi_macro_mln',
            'perc_ricavi_target_su_macro', 'perc_ricavi_categoria', 'perc_attivo_categoria',
            'rnk_naz_ebitda', 'rnk_reg_ebitda', 'rnk_naz_ebit', 'rnk_reg_ebit', 'rnk_naz_prof', 'rnk_reg_prof',
            'rnk_naz_strut1', 'rnk_reg_strut1', 'rnk_naz_strut2', 'rnk_reg_strut2', 'rnk_naz_gear', 'rnk_reg_gear',
            'rnk_naz_rot', 'rnk_reg_rot', 'rnk_naz_cr', 'rnk_reg_cr', 'rnk_naz_qr', 'rnk_reg_qr',
            'med_mg_ebitda', 'med_mg_ebit', 'med_mg_prof', 'med_ind_str1', 'med_ind_str2',
            'med_ind_gear', 'med_ind_cr', 'med_ind_qr', 'med_ind_rot_cap'
        ]
        
        valori_da_sostituire = []
        for k in variabili_sensibili:
            if k in context:
                valore_reale = str(context[k])
                context[k] = testo_censura_sicuro
                if valore_reale not in ["N.D.", "n.d.", "-", "", "0", "0,00", "0.0", "0.00"]:
                    valori_da_sostituire.append(valore_reale)
                
        chiavi_discorsive = ['analisi_trend_ebitda', 'impatto_territoriale']
        for k in chiavi_discorsive:
            if k in context:
                testo_normale = str(context[k])
                for val in valori_da_sostituire:
                    if len(val) >= 2: 
                        testo_normale = testo_normale.replace(val, testo_censura_sicuro)
                context[k] = testo_normale

    # Ora genera il documento in totale sicurezza! (Renderizza il testo crudo)
    doc.render(context)

    # =================================================================
    # 🎨 4. COLORATORE NATIVO CHIRURGICO: Rende ROSSA SOLO la scritta premium
    # =================================================================
    if modalita_teaser:
        from docx.shared import RGBColor

        def colora_paragrafo_chirurgico(p):
            # Interveniamo solo se nel testo globale del paragrafo esiste la scure
            if "🔒 PREMIUM" in p.text:
                testo_completo = p.text
                
                # Ci salviamo lo stile del font del primo run per non perdere la formattazione del template
                font_name = p.runs[0].font.name if p.runs else None
                font_size = p.runs[0].font.size if p.runs else None
                
                # Svuotiamo i runs del paragrafo per ricostruirli in modo pulito
                p.text = "" 
                
                # Tagliamo il testo usando la scritta premium come separatore
                parti = re.split(r'(🔒 PREMIUM)', testo_completo)
                for parte in parti:
                    if not parte:
                        continue
                    
                    # Creiamo un run dedicato per questo frammento di testo
                    run = p.add_run(parte)
                    
                    # Ripristiniamo il font originale per non alterare il layout
                    if font_name: run.font.name = font_name
                    if font_size: run.font.size = font_size
                    
                    # Se questo specifico frammento è la scritta di censura, lo spariamo rosso e bold
                    if "🔒 PREMIUM" in parte:
                        run.font.color.rgb = RGBColor(255, 0, 0)
                        run.font.bold = True

        # 1. Spazzoliamo tutti i paragrafi standard del documento (testi liberi, box, ecc.)
        for p in doc.docx.paragraphs:
            colora_paragrafo_chirurgico(p)

        # 2. Spazzoliamo tutti i paragrafi nascosti dentro le celle delle tabelle
        for table in doc.docx.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        colora_paragrafo_chirurgico(p)

    # =================================================================
    # 📏 FIX SPAZIATURA: Imposta interlinea 1.0 per tutto il testo narrativo
    # =================================================================
    for p in doc.docx.paragraphs:
        if p.text.strip():
            p.paragraph_format.line_spacing = 1.0

    output_word = io.BytesIO()
    doc.save(output_word)
    output_word.seek(0)

    # 🧹 Pulizia del file temporaneo
    try:
        os.remove(template_pulito)
    except:
        pass

    # =================================================================
    # 🔧 POST-PROCESSOR: Unisce i paragrafi frammentati dal rendering
    # =================================================================
    output_word = unisci_paragrafi_frammentati(output_word, ragione_sociale_pulita)

    # =================================================================
    # 🎨 POST-PROCESSOR: Migliora layout, indentazione e struttura
    # =================================================================
    output_word = migliora_layout(output_word)

    return output_word


def unisci_paragrafi_frammentati(output_buffer, ragione_sociale):
    """
    Corregge la frammentazione dei paragrafi causata da variabili inline
    nel template docxtpl. Versione conservativa: non esegue mai merge
    tra stili di paragrafo incompatibili per evitare fusioni errate.
    """
    abbrevs_no_sentence = [
        'S.r.l.', 'S.p.A.', 'S.n.c.', 'S.a.s.', 'S.r.l.s.',
        'n.d.', 'ecc.', 'etc.', 'ca.', 'es.', '1\u00b0.', '2\u00b0.', 'Liv.'
    ]
    placeholder_vals = {'N.D.', 'n.d.', 'N/D', 'n/d', 'N.d.'}
    header_styles = {'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4'}

    def get_text(p):
        return p.text.strip()

    def is_section_header(p):
        style = p.style.name
        text = p.text.strip()
        if style in header_styles:
            return True
        if style == 'Normal' and len(text) < 60 and not any(c in text for c in ['.', ',', ';', ':']):
            return True
        if style == 'Body Text' and len(text) < 50 and not any(c in text for c in ['.', ',', ';', ':', '(', '"', '\'']):
            return True
        # Protect known section titles regardless of style
        known_titles = ['Benchmark Economico', 'Benchmark Patrimoniale', 'Benchmark Finanziario']
        if any(text == title for title in known_titles):
            return True
        return False

    def ends_sentence(text):
        if not text or text[-1] not in ['.', '!', '?']:
            return False
        for a in abbrevs_no_sentence:
            if text.endswith(a):
                return False
        words = text.split()
        if words and re.match(r'^[A-Za-z]{1,3}\.$', words[-1]):
            return False
        return True

    def copy_text_to(target_p, text, prefix=' '):
        r = OxmlElement('w:r')
        if target_p.runs:
            rpr = target_p.runs[0]._r.find(qn('w:rPr'))
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
        t = OxmlElement('w:t')
        t.text = prefix + text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t)
        target_p._p.append(r)

    def remove_p(p):
        """
        Rimuove il paragrafo se non ha forme ancorate.
        Se ha forme ancorate, lo svuota mantenendo la struttura per preservare il layout.
        """
        from docx.oxml.ns import qn as _qn
        wp_ns = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

        # Controlla se ha forme ancorate
        has_anchor = any(
            d.findall('.//' + ('{%s}anchor' % wp_ns))
            for d in p._p.findall('.//' + _qn('w:drawing'))
        )

        if has_anchor:
            # Ha forme ancorate: svuota senza rimuovere
            for child in list(p._p):
                if child.tag not in [_qn('w:pPr')]:
                    p._p.remove(child)
            pPr = p._p.find(_qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                p._p.insert(0, pPr)
            old_spacing = pPr.find(_qn('w:spacing'))
            if old_spacing is not None:
                pPr.remove(old_spacing)
            spacing = OxmlElement('w:spacing')
            spacing.set(_qn('w:before'), '0')
            spacing.set(_qn('w:after'), '0')
            spacing.set(_qn('w:line'), '20')
            spacing.set(_qn('w:lineRule'), 'exact')
            pPr.append(spacing)
            old_rPr = pPr.find(_qn('w:rPr'))
            if old_rPr is not None:
                pPr.remove(old_rPr)
            rPr = OxmlElement('w:rPr')
            sz = OxmlElement('w:sz')
            sz.set(_qn('w:val'), '2')
            szCs = OxmlElement('w:szCs')
            szCs.set(_qn('w:val'), '2')
            rPr.append(sz)
            rPr.append(szCs)
            pPr.append(rPr)
        else:
            # Nessuna forma ancorata: rimuovi completamente
            p._p.getparent().remove(p._p)

    def has_anchor_shape(p):
        """True se il paragrafo ha forme flottanti ancorate — NON rimuovere."""
        wp_ns = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        for drawing in p._p.findall('.//' + qn('w:drawing')):
            if drawing.findall('.//' + ('{%s}anchor' % wp_ns)):
                return True
        return False

    def is_numero(text):
        return bool(re.match(r'^-?\d+[.,]?\d*$', text)) and len(text) < 12

    def is_fragment(text):
        if not text:
            return False
        if text[0] in ['.', ',', '%', ';', ')']:
            return True
        if text[0].islower():
            return True
        if text.startswith('\u2013') or text.startswith('- '):
            return True
        continuators = [
            'si ', 'nel ', 'nella ', 'nelle ', 'una ', 'un ',
            'per ', 'con ', 'che ', 'di ', 'dell', 'nell',
            'al ', 'alla ', 'agli ', 'affronta', 'appare',
            'sono ', 'ha ', 'garantendo', 'riflettendo',
            'evidenziando', 'e la ', 'e il ', 'e i ',
            'risulti ', 'rappresenta', 'conferma', 'indica '
            '- ', 'Anno ', 'mediano ', '2021', '2022', '2023', '2024'
            'N.D. 20', 'N.D. 2021', 'N.D. 2022'
        ]
        return any(text.lower().startswith(c) for c in continuators)

    output_buffer.seek(0)
    doc = docx.Document(output_buffer)

    for _ in range(150):
        changed = False
        ps = list(doc.paragraphs)
        n = len(ps)

        for i in range(1, n - 1):
            curr_text = get_text(ps[i])
            if not curr_text:
                continue  # Skip already-cleared paragraphs

            # Find previous non-empty paragraph
            prev_idx = i - 1
            while prev_idx >= 0 and not get_text(ps[prev_idx]):
                prev_idx -= 1
            if prev_idx < 0:
                continue
            prev_text = get_text(ps[prev_idx])
            if not prev_text:
                continue

            prev_header = is_section_header(ps[prev_idx])

            # Mai processare paragrafi con forme ancorate
            if has_anchor_shape(ps[i]):
                continue

            # CASO 0: Valore placeholder (N.D.) — mai dentro intestazioni
            if curr_text in placeholder_vals and not ends_sentence(prev_text) and not prev_header:
                copy_text_to(ps[prev_idx], curr_text, ' ')
                remove_p(ps[i])
                changed = True
                break

            # CASO 1: Ragione sociale — mai dentro intestazioni
            elif curr_text == ragione_sociale and not prev_header:
                copy_text_to(ps[prev_idx], curr_text, ' ')
                remove_p(ps[i])
                changed = True
                break

            # CASO 2: Numero — solo tra Body Text
            elif (is_numero(curr_text) and not ends_sentence(prev_text) and
                  not prev_header and ps[prev_idx].style.name == 'Body Text'):
                copy_text_to(ps[prev_idx], curr_text, ' ')
                remove_p(ps[i])
                changed = True
                break

            # CASO 3: Punteggiatura — mai dentro intestazioni
            elif curr_text[0] in ['.', ',', '%'] and prev_text and not prev_header:
                if curr_text[0] == '.' and prev_text.endswith('.'):
                    remainder = curr_text[1:].strip()
                    if remainder:
                        copy_text_to(ps[prev_idx], remainder, ' ')
                    remove_p(ps[i])
                else:
                    copy_text_to(ps[prev_idx], curr_text, '')
                    remove_p(ps[i])
                changed = True
                break

            # CASO 4: Frammento — SOLO tra Body Text dello stesso stile
            elif (is_fragment(curr_text) and
                  not ends_sentence(prev_text) and
                  curr_text != ragione_sociale and
                  not prev_header and
                  ps[i].style.name == 'Body Text' and
                  ps[prev_idx].style.name == 'Body Text'):
                copy_text_to(ps[prev_idx], curr_text, ' ')
                remove_p(ps[i])
                changed = True
                break

        if not changed:
            break

    # =============================================================
    # Corregge i doppi punti tra runs adiacenti (es. "N.D." + ".")
    # =============================================================
    for p in doc.paragraphs:
        runs = [r for r in p.runs if r.text]
        for i in range(len(runs) - 1):
            curr_run = runs[i]
            next_run = runs[i + 1]
            # Se il run corrente termina con '.' e il prossimo inizia con '.'
            if curr_run.text.endswith('.') and next_run.text.startswith('.'):
                # Rimuovi il punto iniziale dal run successivo
                next_run.text = next_run.text[1:].lstrip()
                if not next_run.text:
                    next_run._r.getparent().remove(next_run._r)

    result = io.BytesIO()
    doc.save(result)
    result.seek(0)
    return result


def migliora_layout(output_buffer):
    """
    Migliora il layout del documento:
    1. Converte heading vuoti in Normal (pulisce TOC)
    2. Page break prima di ogni Heading 1 con contenuto
    3. KeepWithNext per heading
    4. Indentazione Body Text coerente
    5. Tabelle allineate con il testo, larghezza fissa, colonne uguali
    """
    output_buffer.seek(0)
    doc = docx.Document(output_buffer)

    BODY_INDENT = 412
    PAGE_W = 12240
    LEFT_MARGIN = 720
    RIGHT_MARGIN = 1080
    TEXT_WIDTH = PAGE_W - LEFT_MARGIN - RIGHT_MARGIN   # 10440
    TABLE_WIDTH = TEXT_WIDTH - BODY_INDENT              # 10028

    # ================================================================
    # 1. Heading vuoti → Normal + page break + keepWithNext
    # ================================================================
    first_real_heading_done = False

    for p in doc.paragraphs:
        style = p.style.name
        text = p.text.strip()
        pPr = p._p.find(qn('w:pPr'))

        if style in ['Heading 1', 'Heading 2', 'Heading 3'] and not text:
            try:
                p.style = doc.styles['Normal']
            except:
                pass
            continue

        if style == 'Heading 1' and text:
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                p._p.insert(0, pPr)
            old_pb = pPr.find(qn('w:pageBreakBefore'))
            if old_pb is not None:
                pPr.remove(old_pb)
            if first_real_heading_done:
                pb = OxmlElement('w:pageBreakBefore')
                pb.set(qn('w:val'), 'true')
                pPr.insert(0, pb)
            else:
                first_real_heading_done = True
            kwn = pPr.find(qn('w:keepNext'))
            if kwn is None:
                kwn = OxmlElement('w:keepNext')
                pPr.append(kwn)

        if style in ['Heading 2', 'Heading 3'] and text:
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                p._p.insert(0, pPr)
            kwn = pPr.find(qn('w:keepNext'))
            if kwn is None:
                kwn = OxmlElement('w:keepNext')
                pPr.append(kwn)

   

    # ================================================================
    # 3. Tabelle: larghezza fissa, indent con testo, colonne uguali
    # ================================================================
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    for tbl in doc.element.body.findall('.//' + qn('w:tbl')):
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)

        for tag in [qn('w:tblW'), qn('w:tblInd'), qn('w:jc'), qn('w:tblBorders')]:
            old = tblPr.find(tag)
            if old is not None:
                tblPr.remove(old)

        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:type'), 'dxa')
        tblW.set(qn('w:w'), str(TABLE_WIDTH))
        tblPr.append(tblW)

        tblInd = OxmlElement('w:tblInd')
        tblInd.set(qn('w:type'), 'dxa')
        tblInd.set(qn('w:w'), str(BODY_INDENT))
        tblPr.append(tblInd)

        tblBorders = parse_xml(
            r'<w:tblBorders %s>'
            r'<w:top w:val="single" w:sz="4" w:space="0" w:color="2B3A67"/>'
            r'<w:left w:val="single" w:sz="4" w:space="0" w:color="2B3A67"/>'
            r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="2B3A67"/>'
            r'<w:right w:val="single" w:sz="4" w:space="0" w:color="2B3A67"/>'
            r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="2B3A67"/>'
            r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="2B3A67"/>'
            r'</w:tblBorders>' % nsdecls('w')
        )
        tblPr.append(tblBorders)

        rows = tbl.findall(qn('w:tr'))
        if not rows:
            continue

        num_cols = 0
        for cell in rows[0].findall(qn('w:tc')):
            tcPr = cell.find(qn('w:tcPr'))
            gs = tcPr.find(qn('w:gridSpan')) if tcPr is not None else None
            num_cols += int(gs.get(qn('w:val'), '1')) if gs is not None else 1

        if num_cols < 1:
            continue

        col_width = TABLE_WIDTH // num_cols

        for row in rows:
            for cell in row.findall(qn('w:tc')):
                tcPr = cell.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = OxmlElement('w:tcPr')
                    cell.insert(0, tcPr)
                gs = tcPr.find(qn('w:gridSpan'))
                span = int(gs.get(qn('w:val'), '1')) if gs is not None else 1
                old_tcW = tcPr.find(qn('w:tcW'))
                if old_tcW is not None:
                    tcPr.remove(old_tcW)
                tcW = OxmlElement('w:tcW')
                tcW.set(qn('w:type'), 'dxa')
                tcW.set(qn('w:w'), str(col_width * span))
                tcPr.insert(0, tcW)
    for p in doc.paragraphs:
        if p.style.name == 'Body Text' and p.text.strip():
             for run in p.runs:
                if run.font.size is None or run.font.size >= Pt(12):
                    run.font.size = Pt(11)
  
    # Remove hyperlink character style from TOC entries
   # Remove hyperlinks from TOC by unwrapping w:hyperlink elements
    
    for p in doc.paragraphs:
        if p.style.name in ['TOC 1', 'TOC 2', 'TOC 3']:
            for hyperlink in p._p.findall('.//' + qn('w:hyperlink')):
                parent = hyperlink.getparent()
                idx = list(parent).index(hyperlink)
                for child in list(hyperlink):
                    parent.insert(idx, child)
                    idx += 1
                parent.remove(hyperlink)    
    # Fix inconsistent first line indentation
    for p in doc.paragraphs:
        if p.style.name == 'Body Text' and p.text.strip():
            pPr = p._p.find(qn('w:pPr'))
            if pPr is not None:
                ind = pPr.find(qn('w:ind'))
                if ind is not None:
                    if ind.get(qn('w:firstLine')):
                        ind.attrib.pop(qn('w:firstLine'))
                    if ind.get(qn('w:hanging')):
                        ind.attrib.pop(qn('w:hanging'))            
    for tbl in doc.element.body.findall('.//' + qn('w:tbl')):
        rows = tbl.findall(qn('w:tr'))
        for row in rows:
            cells = row.findall(qn('w:tc'))
            for cell in cells:
                for para in cell.findall('.//' + qn('w:p')):
                    text = ''.join(t.text or '' for t in para.findall('.//' + qn('w:t'))).strip()
                    if re.match(r'^-?[\d\.,°%]+$', text):
                        pPr = para.find(qn('w:pPr'))
                        if pPr is None:
                            pPr = OxmlElement('w:pPr')
                            para.insert(0, pPr)
                        jc = pPr.find(qn('w:jc'))
                        if jc is None:
                            jc = OxmlElement('w:jc')
                            pPr.append(jc)
                        jc.set(qn('w:val'), 'right')
    result = io.BytesIO()
    doc.save(result)
    result.seek(0)
    return result
