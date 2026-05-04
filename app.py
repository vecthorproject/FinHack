import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import re
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell


# ==========================================
# IMPOSTAZIONI PAGINA WEB
# ==========================================
st.set_page_config(page_title="FinHack", layout="wide", page_icon="💹")
st.title("📊 FinHack 😈: Generatore Report")

# ==========================================
# 1. FUNZIONI DEI CAPITOLI (Moduli)
# ==========================================

def elabora_capitolo_1(df_filtered):
    # ==========================================
    # GANCI INIZIALI
    # ==========================================
    df = df_filtered.copy()

    # --- LOGICA PANDAS DEL CAPITOLO 1 ---
    mappatura_fg = {
        'S.R.L. - SRL': 'Società a Responsabilità Limitata (S.r.l)',
        'Società a responsabilità limitata a capitale ridotto': 'Società a Responsabilità Limitata (S.r.l)',
        'S.R.L. semplificata': 'Società a Responsabilità Limitata (S.r.l)',
        'S.R.L. a socio unico - SRLU': 'Società a Responsabilità Limitata (S.r.l)',
        'Società europea - SE': 'Società per Azioni (S.p.A)',
        'Società per azioni - S.P.A. - SPA': 'Società per Azioni (S.p.A)',
        'S.A.P.A. - SAPA': 'Società per Azioni (S.p.A)',
        'S.P.A. a socio unico - SPA': 'Società per Azioni (S.p.A)'
    }

    colonna_fg = 'Forma giuridica nazionale' 
    df['Forma Giuridica Pulita'] = df[colonna_fg].str.replace(r'\s*\(Italia\)', '', regex=True).str.strip()
    df['Macro Forma Giuridica'] = df['Forma Giuridica Pulita'].map(mappatura_fg)
    df_cap1 = df.dropna(subset=['Forma Giuridica Pulita'])

    # FOGLIO "FG"
    fg_detail = df_cap1['Forma Giuridica Pulita'].value_counts().reset_index()
    fg_detail.columns = ['Forma giuridica', 'Conteggio di Ragione socialeCaratteri latini']
    fg_detail = fg_detail.sort_values('Forma giuridica').reset_index(drop=True)
    m_det = len(fg_detail)
    fg_detail.loc[m_det] = ['Totale complessivo', f'=SUM(B2:B{m_det+1})']

    fg_macro = df_cap1['Macro Forma Giuridica'].value_counts().reset_index()
    fg_macro.columns = ['Forma Giuridica', 'V.A.']
    m_mac = len(fg_macro)
    tot_row_excel = 4 + m_mac + 1 
    percents = [f'=F{5+i}/$F${tot_row_excel}*100' for i in range(m_mac)]
    fg_macro['%'] = percents
    fg_macro.loc[m_mac] = ['Totale', f'=SUM(F5:F{tot_row_excel-1})', f'=SUM(G5:G{tot_row_excel-1})']

    # FOGLIO "Liv.Agg. per FG"
    col_attivo = 'Totale Attivo migl EUR 2024'
    col_ricavi = 'Totale valore della produzione migl EUR 2024'

    fin_detail = df_cap1.groupby('Forma Giuridica Pulita')[[col_attivo, col_ricavi]].sum().reset_index()
    fin_detail.columns = ['Etichette di riga', 'Somma di Totale Attivo migl EUR 2024', 'Somma di Totale valore della produzione migl EUR 2024']
    fin_detail['Somma di Totale Attivo migl EUR 2024'] = fin_detail['Somma di Totale Attivo migl EUR 2024'].round(2)
    fin_detail['Somma di Totale valore della produzione migl EUR 2024'] = fin_detail['Somma di Totale valore della produzione migl EUR 2024'].round(2)
    fin_detail = fin_detail.sort_values('Etichette di riga').reset_index(drop=True)
    m_fin_det = len(fin_detail)
    fin_detail.loc[m_fin_det] = ['Totale complessivo', f'=SUM(B2:B{m_fin_det+1})', f'=SUM(C2:C{m_fin_det+1})']

    grouped_fin = df_cap1.groupby('Macro Forma Giuridica')[[col_attivo, col_ricavi]].sum().T
    grouped_fin.index = ['Totale Attivo', 'Totale Ricavi']

    fin_macro_temp = pd.DataFrame(index=grouped_fin.index)
    fin_macro_temp['S.r.l V.A.'] = grouped_fin.get('Società a Responsabilità Limitata (S.r.l)', 0).round(2)
    fin_macro_temp['S.p.A. V.A.'] = grouped_fin.get('Società per Azioni (S.p.A)', 0).round(2)
    fin_macro_temp.reset_index(inplace=True)
    fin_macro_temp.rename(columns={'index': 'Variabile'}, inplace=True)
    fin_macro_temp = fin_macro_temp.iloc[::-1].reset_index(drop=True)

    fin_macro = pd.DataFrame()
    fin_macro['Variabile'] = fin_macro_temp['Variabile']
    fin_macro['S.r.l V.A.'] = fin_macro_temp['S.r.l V.A.']
    
    srl_perc, spa_perc, tot_va, tot_perc = [], [], [], []
    for i in range(len(fin_macro_temp)):
        r = 5 + i 
        tot_va.append(f'=F{r}+H{r}')
        srl_perc.append(f'=IF(J{r}=0,0,F{r}/J{r}*100)')
        spa_perc.append(f'=IF(J{r}=0,0,H{r}/J{r}*100)')
        tot_perc.append(f'=G{r}+I{r}')

    fin_macro['S.r.l %'] = srl_perc
    fin_macro['S.p.A. V.A.'] = fin_macro_temp['S.p.A. V.A.']
    fin_macro['S.p.A. %'] = spa_perc
    fin_macro['Totale V.A.'] = tot_va
    fin_macro['Totale %'] = tot_perc


    # ==========================================
    # GANCIO DI MEZZO (Scrive in RAM invece che su disco)
    # ==========================================
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        fg_detail.to_excel(writer, sheet_name='FG', index=False, startcol=0, startrow=0)
        fg_macro.to_excel(writer, sheet_name='FG', index=False, startcol=4, startrow=3)
        fin_detail.to_excel(writer, sheet_name='Liv.Agg. per FG', index=False, startcol=0, startrow=0)
        fin_macro.to_excel(writer, sheet_name='Liv.Agg. per FG', index=False, startcol=4, startrow=3)


    # ==========================================
    # GANCIO DI FORMATTAZIONE
    # ==========================================
    output_buffer.seek(0)
    wb = load_workbook(output_buffer)

    header_fill = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    alt_row_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    total_font = Font(bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    def format_table(worksheet, start_row, start_col, dataframe, is_count=False, has_total_row=True):
        end_row = start_row + len(dataframe)
        end_col = start_col + len(dataframe.columns) - 1
        
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row, column=col)
                
                if row == start_row:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.border = thin_border
                    if row % 2 != 0:
                        cell.fill = alt_row_fill
                
                if has_total_row and row == end_row:
                    cell.font = total_font
                    cell.fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')

                if row > start_row and col > start_col:
                    header_val = str(worksheet.cell(row=start_row, column=col).value)
                    if "%" in header_val:
                        cell.number_format = '0.00'
                    elif is_count:
                        cell.number_format = '#,##0'
                    else:
                        cell.number_format = '#,##0.00'

    ws_fg = wb['FG']
    format_table(ws_fg, 1, 1, fg_detail, is_count=True, has_total_row=True)
    format_table(ws_fg, 4, 5, fg_macro, is_count=True, has_total_row=True)

    ws_fin = wb['Liv.Agg. per FG']
    format_table(ws_fin, 1, 1, fin_detail, is_count=False, has_total_row=True)
    format_table(ws_fin, 4, 5, fin_macro, is_count=False, has_total_row=False) 

    for ws in [ws_fg, ws_fin]:
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[column].width = min((max_length + 2), 50)


    # ==========================================
    # GANCI FINALI
    # ==========================================
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    
    return final_output



def elabora_capitolo_2(df_filtered):
    import io
    import pandas as pd

    # ==========================================
    # GANCI INIZIALI
    # ==========================================
    df_base = df_filtered.copy()

    # --- LOGICA PANDAS ---
    cols = [
        'Ragione socialeCaratteri latini', 'NUTS2', 'Totale Attivo migl EUR 2024', 
        'Totale valore della produzione migl EUR 2024', 'Numero dipendenti 2024'
    ]
    # Filtriamo solo le colonne che esistono realmente nel file
    cols_to_use = [c for c in cols if c in df_base.columns]
    df_base = df_base[cols_to_use].copy()
    
    df_base.rename(columns={
        'NUTS2': 'Regione', 
        'Totale valore della produzione migl EUR 2024': 'Totale Ricavi migl EUR 2024',
        'Ragione socialeCaratteri latini': 'Ragione Sociale'
    }, inplace=True)

    def pulisci_regione(nome_regione):
        if pd.isna(nome_regione): return 'Altro'
        if ' - ' in str(nome_regione): return str(nome_regione).split(' - ', 1)[1]
        return str(nome_regione)
        
    if 'Regione' in df_base.columns:
        df_base['Nome Regione'] = df_base['Regione'].apply(pulisci_regione)
    else:
        df_base['Nome Regione'] = 'Altro'

    for col in ['Totale Attivo migl EUR 2024', 'Totale Ricavi migl EUR 2024']:
        if col in df_base.columns:
            df_base[col] = pd.to_numeric(df_base[col], errors='coerce')
            
    if 'Numero dipendenti 2024' in df_base.columns:
        df_base['Numero dipendenti 2024'] = pd.to_numeric(df_base['Numero dipendenti 2024'], errors='coerce')

    def get_macro(nuts2):
        if pd.isna(nuts2): return 'Altro'
        code = str(nuts2)[:3]
        if code == 'ITC': return 'Nord Ovest'
        elif code == 'ITH': return 'Nord Est'
        elif code == 'ITI': return 'Centro'
        elif code in ['ITF', 'ITG']: return 'Sud e Isole'
        else: return 'Altro'

    if 'Regione' in df_base.columns:
        df_base['Macroregione'] = df_base['Regione'].apply(get_macro)
    else:
         df_base['Macroregione'] = 'Altro'

    pivot_reg = df_base.groupby(['Macroregione', 'Nome Regione']).agg({
        'Ragione Sociale': 'count',
        'Totale Ricavi migl EUR 2024': 'sum',
        'Totale Attivo migl EUR 2024': 'sum',
        'Numero dipendenti 2024': 'sum'
    }).rename(columns={'Ragione Sociale': 'Imprese'})

    pivot_reg['Totale Ricavi migl EUR 2024'] = pivot_reg['Totale Ricavi migl EUR 2024'].round(2)
    pivot_reg['Totale Attivo migl EUR 2024'] = pivot_reg['Totale Attivo migl EUR 2024'].round(2)
    pivot_reg['Numero dipendenti 2024'] = pivot_reg['Numero dipendenti 2024'].fillna(0).astype(int)
    pivot_reg['Imprese'] = pivot_reg['Imprese'].fillna(0).astype(int)


    # ==========================================
    # GANCIO DI MEZZO E FORMATTAZIONE CON XLSXWRITER
    # ==========================================
    output_buffer = io.BytesIO()
    
    writer = pd.ExcelWriter(output_buffer, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet('Dati Aggregate')

    # Palette Stili
    format_header_blue = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#4F81BD', 'font_color': 'white', 'border': 1})
    format_subheader = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D9D9D9', 'border': 1})
    format_regione = workbook.add_format({'border': 1})
    format_macro_bold = workbook.add_format({'bold': True, 'fg_color': '#F2F2F2', 'border': 1})
    format_italia = workbook.add_format({'bold': True, 'fg_color': '#D9E1F2', 'top': 1, 'bottom': 6})
    
    f_int = workbook.add_format({'num_format': '#,##0', 'border': 1})
    f_dec = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
    f_perc = workbook.add_format({'num_format': '0.00%', 'border': 1})
    
    f_macro_int = workbook.add_format({'bold': True, 'num_format': '#,##0', 'fg_color': '#F2F2F2', 'border': 1})
    f_macro_dec = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'fg_color': '#F2F2F2', 'border': 1})
    f_macro_perc = workbook.add_format({'bold': True, 'num_format': '0.00%', 'fg_color': '#F2F2F2', 'border': 1})

    f_ita_int = workbook.add_format({'bold': True, 'num_format': '#,##0', 'fg_color': '#D9E1F2', 'top': 1, 'bottom': 6})
    f_ita_dec = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'fg_color': '#D9E1F2', 'top': 1, 'bottom': 6})
    f_ita_perc = workbook.add_format({'bold': True, 'num_format': '0.00%', 'fg_color': '#D9E1F2', 'top': 1, 'bottom': 6})

    # Intestazioni unite
    worksheet.write(0, 0, 'Regioni', format_header_blue)
    worksheet.merge_range(0, 1, 0, 2, 'Imprese', format_header_blue)
    worksheet.merge_range(0, 3, 0, 4, 'Ricavi', format_header_blue)
    worksheet.merge_range(0, 5, 0, 6, 'Totale Attivo', format_header_blue)
    worksheet.merge_range(0, 7, 0, 8, 'Dipendenti', format_header_blue)
    
    subheaders = ['', 'V.A.', '%', 'V.A.', '%', 'V.A.', '%', 'V.A.', '%']
    for col_num, sh in enumerate(subheaders):
        if col_num > 0: worksheet.write(1, col_num, sh, format_subheader)
        else: worksheet.write(1, col_num, sh)

    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:I', 15)

    macros_order = ['Nord Ovest', 'Nord Est', 'Centro', 'Sud e Isole']
    macros_present = [m for m in macros_order if m in pivot_reg.index.get_level_values(0)]
    
    num_regions = len(pivot_reg)
    num_macros = len(macros_present)
    riga_italia_idx = 2 + num_regions + num_macros
    riga_italia_excel = riga_italia_idx + 1

    current_idx = 2
    macro_tot_rows = []

    for macro in macros_present:
        df_macro = pivot_reg.loc[macro]
        start_reg_idx = current_idx
        for reg, row_data in df_macro.iterrows():
            worksheet.write(current_idx, 0, reg, format_regione)
            worksheet.write(current_idx, 1, row_data['Imprese'], f_int)
            worksheet.write(current_idx, 3, row_data['Totale Ricavi migl EUR 2024'], f_dec)
            worksheet.write(current_idx, 5, row_data['Totale Attivo migl EUR 2024'], f_dec)
            worksheet.write(current_idx, 7, row_data['Numero dipendenti 2024'], f_int)
            for c, v_col in zip([2, 4, 6, 8], ['B', 'D', 'F', 'H']):
                worksheet.write_formula(current_idx, c, f"={v_col}{current_idx+1}/{v_col}${riga_italia_excel}", f_perc)
            current_idx += 1
        end_reg_idx = current_idx - 1

        worksheet.write(current_idx, 0, macro, format_macro_bold)
        for c, v_col in zip([1, 3, 5, 7], ['B', 'D', 'F', 'H']):
            worksheet.write_formula(current_idx, c, f"=SUM({v_col}{start_reg_idx+1}:{v_col}{end_reg_idx+1})", f_macro_int if c in [1,7] else f_macro_dec)
        for c, v_col in zip([2, 4, 6, 8], ['B', 'D', 'F', 'H']):
            worksheet.write_formula(current_idx, c, f"={v_col}{current_idx+1}/{v_col}${riga_italia_excel}", f_macro_perc)
        macro_tot_rows.append(current_idx)
        current_idx += 1

    worksheet.write(current_idx, 0, "Italia", format_italia)
    for c, v_col in zip([1, 3, 5, 7], ['B', 'D', 'F', 'H']):
        if macro_tot_rows:
            sum_str = "+".join([f"{v_col}{r+1}" for r in macro_tot_rows])
            worksheet.write_formula(current_idx, c, f"={sum_str}", f_ita_int if c in [1,7] else f_ita_dec)
    for c, v_col in zip([2, 4, 6, 8], ['C', 'E', 'G', 'I']):
        if macro_tot_rows:
            sum_str = "+".join([f"{v_col}{r+1}" for r in macro_tot_rows])
            worksheet.write_formula(current_idx, c, f"={sum_str}", f_ita_perc)

    col_macro_start = 11 
    worksheet.write(1, col_macro_start, 'Macroregione', format_header_blue)
    worksheet.write(1, col_macro_start+1, 'Imprese %', format_header_blue)
    worksheet.write(1, col_macro_start+2, 'Ricavi %', format_header_blue)
    worksheet.write(1, col_macro_start+3, 'Attivo %', format_header_blue)
    worksheet.write(1, col_macro_start+4, 'Dipendenti %', format_header_blue)

    for i, (macro, m_row) in enumerate(zip(macros_present, macro_tot_rows)):
        r = 2 + i
        worksheet.write(r, col_macro_start, macro)
        worksheet.write_formula(r, col_macro_start+1, f"=C{m_row+1}", f_perc)
        worksheet.write_formula(r, col_macro_start+2, f"=E{m_row+1}", f_perc)
        worksheet.write_formula(r, col_macro_start+3, f"=G{m_row+1}", f_perc)
        worksheet.write_formula(r, col_macro_start+4, f"=I{m_row+1}", f_perc)

    col_reg_start = 17 
    worksheet.write(1, col_reg_start, 'Regione', format_header_blue)
    worksheet.write(1, col_reg_start+1, 'Imprese', format_header_blue)
    worksheet.write(1, col_reg_start+2, 'Ricavi', format_header_blue)
    worksheet.write(1, col_reg_start+3, 'Attivo', format_header_blue)
    worksheet.write(1, col_reg_start+4, 'Dipendenti', format_header_blue)

    reg_idx = 2
    for r in range(2, riga_italia_idx):
        if r not in macro_tot_rows: 
            worksheet.write_formula(reg_idx, col_reg_start, f"=A{r+1}")
            worksheet.write_formula(reg_idx, col_reg_start+1, f"=B{r+1}", f_int)
            worksheet.write_formula(reg_idx, col_reg_start+2, f"=D{r+1}", f_dec)
            worksheet.write_formula(reg_idx, col_reg_start+3, f"=F{r+1}", f_dec)
            worksheet.write_formula(reg_idx, col_reg_start+4, f"=H{r+1}", f_int)
            reg_idx += 1

    worksheet.set_column('L:V', 14)

    def crea_torta(titolo, colonna_valori, pos_cella):
        chart = workbook.add_chart({'type': 'pie'})
        chart.add_series({
            'name':       titolo,
            'categories': ['Dati Aggregate', 2, col_macro_start, 2 + num_macros - 1, col_macro_start], 
            'values':     ['Dati Aggregate', 2, colonna_valori, 2 + num_macros - 1, colonna_valori],
            'data_labels': {'percentage': True, 'position': 'outside_end'}
        })
        chart.set_title({'name': titolo})
        chart.set_style(10)
        chart.set_size({'width': 400, 'height': 280}) 
        worksheet.insert_chart(pos_cella, chart) 

    def crea_istogramma(titolo, colonna_valori, pos_cella):
        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name':       titolo,
            'categories': ['Dati Aggregate', 2, col_reg_start, reg_idx - 1, col_reg_start], 
            'values':     ['Dati Aggregate', 2, colonna_valori, reg_idx - 1, colonna_valori],
        })
        chart.set_title({'name': titolo})
        chart.set_legend({'none': True})
        chart.set_style(10)
        chart.set_size({'width': 550, 'height': 280}) 
        worksheet.insert_chart(pos_cella, chart) 

    crea_torta('Ripartizione Imprese (%)', col_macro_start+1, 'L9')
    crea_torta('Ripartizione Ricavi (%)', col_macro_start+2, 'L24')
    crea_torta('Ripartizione Attivo (%)', col_macro_start+3, 'L39')
    crea_torta('Ripartizione Dipendenti (%)', col_macro_start+4, 'L54')

    col_chart_istogrammi = 'X'
    crea_istogramma('Imprese per Regione', col_reg_start+1, f'{col_chart_istogrammi}2')
    crea_istogramma('Ricavi per Regione (migl EUR)', col_reg_start+2, f'{col_chart_istogrammi}17')
    crea_istogramma('Attivo per Regione (migl EUR)', col_reg_start+3, f'{col_chart_istogrammi}32')
    crea_istogramma('Dipendenti per Regione', col_reg_start+4, f'{col_chart_istogrammi}47')

    ws_quartili = workbook.add_worksheet('Quartili')
    
    df_raw = df_base[['Totale Attivo migl EUR 2024', 'Totale Ricavi migl EUR 2024']].dropna()
    df_raw = df_raw.sort_values(by='Totale Ricavi migl EUR 2024', ascending=False)
    
    ws_quartili.write(0, 0, 'Totale Attivo migl EUR 2024', format_header_blue)
    ws_quartili.write(0, 1, 'Totale valore della produzione migl EUR 2024', format_header_blue)
    
    for r_idx, (_, row) in enumerate(df_raw.iterrows(), 1):
        ws_quartili.write(r_idx, 0, row['Totale Attivo migl EUR 2024'], f_dec)
        ws_quartili.write(r_idx, 1, row['Totale Ricavi migl EUR 2024'], f_dec)

    ws_quartili.set_column('A:B', 30)
    ultima_riga_dati = len(df_raw) + 1

    ws_quartili.write(0, 3, 'Variabile', format_header_blue)
    ws_quartili.write(0, 4, 'V.A.', format_header_blue)
    ws_quartili.write(1, 3, 'Totale Ricavi')
    ws_quartili.write_formula(1, 4, f"=SUM(B2:B{ultima_riga_dati})", f_dec)
    ws_quartili.write(2, 3, 'Totale Attivo')
    ws_quartili.write_formula(2, 4, f"=SUM(A2:A{ultima_riga_dati})", f_dec)

    ws_quartili.set_column('D:D', 15)
    ws_quartili.set_column('E:E', 20)

    ws_quartili.write(0, 6, 'Quartile', format_header_blue)
    ws_quartili.write(0, 7, 'Totale Attivo', format_header_blue)
    ws_quartili.write(0, 8, 'Totale Ricavi', format_header_blue)

    nomi_quartili = ['Minimo', '1°', '2°', '3°', '4°']
    for i, nome in enumerate(nomi_quartili):
        riga = i + 1
        ws_quartili.write(riga, 6, nome, format_subheader)
        ws_quartili.write_formula(riga, 7, f"=ROUND(QUARTILE(A2:A{ultima_riga_dati}, {i}), 2)", f_dec)
        ws_quartili.write_formula(riga, 8, f"=ROUND(QUARTILE(B2:B{ultima_riga_dati}, {i}), 2)", f_dec)

    start_r = 8
    ws_quartili.write(start_r, 6, 'Quartile', format_header_blue)
    ws_quartili.write(start_r, 7, 'Totale Attivo', format_header_blue)
    ws_quartili.write(start_r, 8, 'Totale Ricavi', format_header_blue)

    intervalli_nomi = ['1°', '2°', '3°', '4°']
    for i in range(4):
        par_chiusa = ')' if i < 3 else ']'
        riga_min = i + 2 
        riga_max = i + 3 
        formula_attivo = f'="[" & TEXT(H{riga_min}, "#.##0,00") & " - " & TEXT(H{riga_max}, "#.##0,00") & "{par_chiusa}"'
        formula_ricavi = f'="[" & TEXT(I{riga_min}, "#.##0,00") & " - " & TEXT(I{riga_max}, "#.##0,00") & "{par_chiusa}"'
        
        ws_quartili.write(start_r + 1 + i, 6, intervalli_nomi[i], format_subheader)
        ws_quartili.write_formula(start_r + 1 + i, 7, formula_attivo)
        ws_quartili.write_formula(start_r + 1 + i, 8, formula_ricavi)

    ws_quartili.set_column('G:I', 35)

    # ==========================================
    # GANCI FINALI
    # ==========================================
    writer.close()
    output_buffer.seek(0)
    
    return output_buffer


def elabora_capitolo_3(df_filtered):
    import io
    import pandas as pd
    import numpy as np
    import xlsxwriter
    from xlsxwriter.utility import xl_rowcol_to_cell
    import re

    # Funzione interna che hai creato tu, adattata per girare qui dentro
    def costruisci_sezione_analisi(writer, workbook, df_raw, formati, keyword_ricerca, sheet_data, sheet_stats, chart_title, y_axis_name, rename_dict):
        # 1. Filtro Colonne per la sezione corrente
        base_cols = [c for c in df_raw.columns if 'ragione' in str(c).lower() or 'bvd' in str(c).lower()]
        metric_cols = [c for c in df_raw.columns if keyword_ricerca in str(c).lower()]
        
        if not metric_cols:
            return

        df = df_raw[base_cols + metric_cols].copy()
        df = df.replace(['n.d.', 'n.a.', 'n.s.', ''], np.nan)
        num_rows = len(df)

        # 2. Creazione Fogli
        df.to_excel(writer, sheet_name=sheet_data, index=False, startrow=0)
        worksheet_data = writer.sheets[sheet_data]
        worksheet_stats = workbook.add_worksheet(sheet_stats)

        # 3. Compilazione Foglio 1: Dati e Autofit
        for i, col in enumerate(df.columns):
            worksheet_data.write(0, i, col, formati['header'])
            col_data = df[col].dropna()
            max_len = len(str(col)) + 2 if col_data.empty else max(col_data.astype(str).map(len).max(), len(str(col))) + 2 
            worksheet_data.set_column(i, i, min(max_len, 45))
            for row in range(1, num_rows + 1):
                val = df.iat[row-1, i]
                if pd.isna(val):
                    worksheet_data.write(row, i, "", formati['data_text'])
                elif isinstance(val, (int, float)):
                    worksheet_data.write_number(row, i, val, formati['data_num'])
                else:
                    worksheet_data.write_string(row, i, str(val), formati['data_text'])

        # 4. Preparazione Metriche per il Foglio 2
        numeric_cols_idx = [df.columns.get_loc(c) for c in metric_cols]
        metrics_dict = {} 
        
        for col_idx in numeric_cols_idx:
            col_name = df.columns[col_idx]
            anno_match = re.search(r'\d{4}', col_name)
            if anno_match:
                anno = int(anno_match.group(0))
                metric_base = re.sub(r'\d{4}', '', col_name).replace('(*) %', '').replace('migl EUR', '').replace('  ', ' ').strip()
                for k, v in rename_dict.items():
                    if k.lower() in metric_base.lower():
                        metric_base = v
                
                if metric_base not in metrics_dict: metrics_dict[metric_base] = {}
                metrics_dict[metric_base][anno] = col_idx

        all_years = sorted(list(set(anno for m in metrics_dict for anno in metrics_dict[m].keys())))
        metrics_list = list(metrics_dict.keys())

        col_left = 0
        col_mid = len(all_years) + 2
        col_right = col_mid + (len(all_years)*2 - 1) + 2

        worksheet_stats.set_column(col_left, col_left, 22)
        worksheet_stats.set_column(col_left + 1, col_left + len(all_years), 12)
        worksheet_stats.set_column(col_mid, col_mid, 22)
        worksheet_stats.set_column(col_mid + 1, col_mid + (len(all_years)*2), 10)
        worksheet_stats.set_column(col_right, col_right, 22)
        worksheet_stats.set_column(col_right + 1, col_right + len(all_years), 12)

        median_cells = {}
        current_row = 1
        stats_formulas = [('Media', 'AVERAGE'), ('Mediana', 'MEDIAN'), ('Asimmetria', 'SKEW'), ('Curtosi', 'KURT'), ('Deviazione standard', 'STDEV')]

        # 5. Costruzione Blocco Sinistro (Stats) e Centrale (YoY)
        for metric in metrics_list:
            median_cells[metric] = {}
            
            # --- BLOCCO SINISTRO ---
            worksheet_stats.write(current_row, col_left, "", formati['header'])
            for i, year in enumerate(all_years):
                worksheet_stats.write(current_row, col_left + 1 + i, year, formati['header'])
                
            for s_idx, (stat_name, func) in enumerate(stats_formulas):
                r = current_row + 1 + s_idx
                worksheet_stats.write(r, col_left, stat_name, formati['label'])
                for i, year in enumerate(all_years):
                    if year in metrics_dict[metric]:
                        data_col_idx = metrics_dict[metric][year]
                        col_letter = xlsxwriter.utility.xl_col_to_name(data_col_idx)
                        data_range = f"'{sheet_data}'!{col_letter}2:{col_letter}{num_rows+1}"
                        formula = f"={func}({data_range})"
                        worksheet_stats.write_formula(r, col_left + 1 + i, formula, formati['data_num'])
                        if stat_name == 'Mediana':
                            median_cells[metric][year] = xl_rowcol_to_cell(r, col_left + 1 + i)
                    else:
                        worksheet_stats.write(r, col_left + 1 + i, "n.d.", formati['data_text'])

            # --- BLOCCO CENTRALE ---
            c_offset = 1
            for i, year in enumerate(all_years):
                if i == 0:
                    worksheet_stats.write(current_row, col_mid + c_offset, year, formati['header'])
                    worksheet_stats.write(current_row + 1, col_mid + c_offset, "V.m.", formati['header'])
                    c_offset += 1
                else:
                    worksheet_stats.merge_range(current_row, col_mid + c_offset, current_row, col_mid + c_offset + 1, year, formati['header'])
                    worksheet_stats.write(current_row + 1, col_mid + c_offset, "V.m.", formati['header'])
                    worksheet_stats.write(current_row + 1, col_mid + c_offset + 1, "Δ %", formati['header'])
                    c_offset += 2
                    
            r_data = current_row + 2
            worksheet_stats.write(r_data, col_mid, metric, formati['label'])
            c_offset = 1
            prev_target_cell = None
            for i, year in enumerate(all_years):
                if year in median_cells[metric]:
                    source_cell = median_cells[metric][year]
                    curr_target_cell = xl_rowcol_to_cell(r_data, col_mid + c_offset)
                    worksheet_stats.write_formula(r_data, col_mid + c_offset, f"={source_cell}", formati['data_num'])
                    
                    if i > 0:
                        if prev_target_cell:
                            delta_formula = f"=IFERROR(({curr_target_cell}-{prev_target_cell})/ABS({prev_target_cell}), 0)"
                            worksheet_stats.write_formula(r_data, col_mid + c_offset + 1, delta_formula, formati['pct'])
                        else:
                            worksheet_stats.write(r_data, col_mid + c_offset + 1, "n.d.", formati['data_text'])
                        prev_target_cell = curr_target_cell
                        c_offset += 2
                    else:
                        prev_target_cell = curr_target_cell
                        c_offset += 1

            current_row += 8 

        # 6. Costruzione Blocco Destro (Riassunto Consolidato)
        r_right = 1
        worksheet_stats.write(r_right, col_right, "", formati['header'])
        for i, year in enumerate(all_years):
            worksheet_stats.write(r_right, col_right + 1 + i, year, formati['header'])
            
        for m_idx, metric in enumerate(metrics_list):
            r_right_data = r_right + 1 + m_idx
            worksheet_stats.write(r_right_data, col_right, metric, formati['label'])
            for i, year in enumerate(all_years):
                if year in median_cells[metric]:
                    source_cell = median_cells[metric][year]
                    worksheet_stats.write_formula(r_right_data, col_right + 1 + i, f"={source_cell}", formati['data_num'])
                else:
                    worksheet_stats.write(r_right_data, col_right + 1 + i, "n.d.", formati['data_text'])

        # 7. Inserimento Grafico Dinamico
        if metrics_list:
            chart = workbook.add_chart({'type': 'column'})
            for m_idx, metric in enumerate(metrics_list):
                r_idx = r_right + 1 + m_idx
                start_col = xlsxwriter.utility.xl_col_to_name(col_right + 1)
                end_col = xlsxwriter.utility.xl_col_to_name(col_right + len(all_years))
                
                chart.add_series({
                    'name': f"='{sheet_stats}'!${xlsxwriter.utility.xl_col_to_name(col_right)}${r_idx + 1}",
                    'categories': f"='{sheet_stats}'!${start_col}$2:${end_col}$2",
                    'values': f"='{sheet_stats}'!${start_col}${r_idx + 1}:${end_col}${r_idx + 1}",
                    'data_labels': {'value': True, 'num_format': '#,##0.##'}
                })
                
            chart.set_legend({'position': 'bottom'})
            chart.set_style(11)
            chart.set_y_axis({'name': y_axis_name, 'major_gridlines': {'visible': True}})
            chart.set_title({'name': chart_title})
            chart.set_size({'width': 600, 'height': 350})

            chart_row = r_right + len(metrics_list) + 3 
            worksheet_stats.insert_chart(chart_row, col_right, chart)

        worksheet_stats.ignore_errors({'formula_differs': 'G1:Z500', 'number_stored_as_text': 'A1:Z500'})

    # ==========================================
    # LOGICA DI ESECUZIONE 
    # ==========================================
    df_raw = df_filtered.copy()
    output_buffer = io.BytesIO()
    
    writer = pd.ExcelWriter(output_buffer, engine='xlsxwriter')
    workbook = writer.book

    formati = {
        'header': workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
        'label': workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'left', 'valign': 'vcenter'}),
        'data_num': workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'align': 'right', 'valign': 'vcenter'}),
        'data_text': workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'}),
        'pct': workbook.add_format({'border': 1, 'num_format': '0.00%', 'align': 'right', 'valign': 'vcenter'})
    }

    costruisci_sezione_analisi(
        writer, workbook, df_raw, formati,
        keyword_ricerca='margin',
        sheet_data='3a. Dati Eq. Economico',
        sheet_stats='3b. Stat. Eq. Economico',
        chart_title='Andamento Mediano Margini',
        y_axis_name='Percentuale (%)',
        rename_dict={} 
    )

    costruisci_sezione_analisi(
        writer, workbook, df_raw, formati,
        keyword_ricerca='produzione', 
        sheet_data='4a. Dati Svil. Dimensionale',
        sheet_stats='4b. Stat. Svil. Dimensionale',
        chart_title='Andamento Mediano Ricavi',
        y_axis_name='Migliaia di Euro (€)',
        rename_dict={'produzione': 'Ricavi'} 
    )

    writer.close()
    output_buffer.seek(0)
    
    return output_buffer


def elabora_capitolo_4(df_filtered):
    import io
    import pandas as pd
    import numpy as np
    import xlsxwriter
    from xlsxwriter.utility import xl_rowcol_to_cell
    import re

    def costruisci_sezione_analisi(writer, workbook, df_raw, formati, keyword_ricerca, sheet_data, sheet_stats, chart_title, y_axis_name, rename_dict):
        # 1. Filtro Colonne
        base_cols = [c for c in df_raw.columns if 'ragione' in str(c).lower() or 'bvd' in str(c).lower()]
        
        if isinstance(keyword_ricerca, list):
            metric_cols = [c for c in df_raw.columns if any(kw.lower() in str(c).lower() for kw in keyword_ricerca)]
        else:
            metric_cols = [c for c in df_raw.columns if keyword_ricerca.lower() in str(c).lower()]

        if not metric_cols:
            return

        df = df_raw[base_cols + metric_cols].copy()
        df = df.replace(['n.d.', 'n.a.', 'n.s.', ''], np.nan)
        num_rows = len(df)

        # 2. Creazione Fogli
        df.to_excel(writer, sheet_name=sheet_data, index=False, startrow=0)
        worksheet_data = writer.sheets[sheet_data]
        worksheet_stats = workbook.add_worksheet(sheet_stats)

        # 3. Compilazione Foglio 1: Dati e Autofit
        for i, col in enumerate(df.columns):
            worksheet_data.write(0, i, col, formati['header'])
            col_data = df[col].dropna()
            max_len = len(str(col)) + 2 if col_data.empty else max(col_data.astype(str).map(len).max(), len(str(col))) + 2
            worksheet_data.set_column(i, i, min(max_len, 45))
            for row in range(1, num_rows + 1):
                val = df.iat[row-1, i]
                if pd.isna(val):
                    worksheet_data.write(row, i, "", formati['data_text'])
                elif isinstance(val, (int, float)):
                    worksheet_data.write_number(row, i, val, formati['data_num'])
                else:
                    worksheet_data.write_string(row, i, str(val), formati['data_text'])

        # 4. Preparazione Metriche per il Foglio 2
        numeric_cols_idx = [df.columns.get_loc(c) for c in metric_cols]
        metrics_dict = {}

        for col_idx in numeric_cols_idx:
            col_name = df.columns[col_idx]
            anno_match = re.search(r'\d{4}', col_name)
            if anno_match:
                anno = int(anno_match.group(0))
                metric_base = re.sub(r'\d{4}', '', col_name).replace('(*) %', '').replace('(*)', '').replace('migl EUR', '').replace(' ', ' ').strip()
                
                for k, v in rename_dict.items():
                    if k.lower() in metric_base.lower():
                        metric_base = v

                if metric_base not in metrics_dict: metrics_dict[metric_base] = {}
                metrics_dict[metric_base][anno] = col_idx

        all_years = sorted(list(set(anno for m in metrics_dict for anno in metrics_dict[m].keys())))
        metrics_list = list(metrics_dict.keys())

        col_left = 0
        col_mid = len(all_years) + 2
        col_right = col_mid + (len(all_years)*2 - 1) + 2

        worksheet_stats.set_column(col_left, col_left, 22)
        worksheet_stats.set_column(col_left + 1, col_left + len(all_years), 12)
        worksheet_stats.set_column(col_mid, col_mid, 22)
        worksheet_stats.set_column(col_mid + 1, col_mid + (len(all_years)*2), 10)
        worksheet_stats.set_column(col_right, col_right, 22)
        worksheet_stats.set_column(col_right + 1, col_right + len(all_years), 12)

        median_cells = {}
        current_row = 1
        stats_formulas = [('Media', 'AVERAGE'), ('Mediana', 'MEDIAN'), ('Asimmetria', 'SKEW'), ('Curtosi', 'KURT'), ('Deviazione standard', 'STDEV')]

        # 5. Costruzione Blocco Sinistro (Stats) e Centrale (YoY)
        for metric in metrics_list:
            median_cells[metric] = {}

            worksheet_stats.write(current_row, col_left, "", formati['header'])
            for i, year in enumerate(all_years):
                worksheet_stats.write(current_row, col_left + 1 + i, year, formati['header'])

            for s_idx, (stat_name, func) in enumerate(stats_formulas):
                r = current_row + 1 + s_idx
                worksheet_stats.write(r, col_left, stat_name, formati['label'])
                for i, year in enumerate(all_years):
                    if year in metrics_dict[metric]:
                        data_col_idx = metrics_dict[metric][year]
                        col_letter = xlsxwriter.utility.xl_col_to_name(data_col_idx)
                        data_range = f"'{sheet_data}'!{col_letter}2:{col_letter}{num_rows+1}"
                        formula = f"={func}({data_range})"
                        worksheet_stats.write_formula(r, col_left + 1 + i, formula, formati['data_num'])
                        if stat_name == 'Mediana':
                            median_cells[metric][year] = xl_rowcol_to_cell(r, col_left + 1 + i)
                    else:
                        worksheet_stats.write(r, col_left + 1 + i, "n.d.", formati['data_text'])

            c_offset = 1
            for i, year in enumerate(all_years):
                if i == 0:
                    worksheet_stats.write(current_row, col_mid + c_offset, year, formati['header'])
                    worksheet_stats.write(current_row + 1, col_mid + c_offset, "V.m.", formati['header'])
                    c_offset += 1
                else:
                    worksheet_stats.merge_range(current_row, col_mid + c_offset, current_row, col_mid + c_offset + 1, year, formati['header'])
                    worksheet_stats.write(current_row + 1, col_mid + c_offset, "V.m.", formati['header'])
                    worksheet_stats.write(current_row + 1, col_mid + c_offset + 1, "Δ %", formati['header'])
                    c_offset += 2

            r_data = current_row + 2
            worksheet_stats.write(r_data, col_mid, metric, formati['label'])
            c_offset = 1
            prev_target_cell = None
            for i, year in enumerate(all_years):
                if year in median_cells[metric]:
                    source_cell = median_cells[metric][year]
                    curr_target_cell = xl_rowcol_to_cell(r_data, col_mid + c_offset)
                    worksheet_stats.write_formula(r_data, col_mid + c_offset, f"={source_cell}", formati['data_num'])

                    if i > 0:
                        if prev_target_cell:
                            delta_formula = f"=IFERROR(({curr_target_cell}-{prev_target_cell})/ABS({prev_target_cell}), 0)"
                            worksheet_stats.write_formula(r_data, col_mid + c_offset + 1, delta_formula, formati['pct'])
                        else:
                            worksheet_stats.write(r_data, col_mid + c_offset + 1, "n.d.", formati['data_text'])
                        prev_target_cell = curr_target_cell
                        c_offset += 2
                    else:
                        prev_target_cell = curr_target_cell
                        c_offset += 1
                else:
                    prev_target_cell = curr_target_cell
                    c_offset += 1

            current_row += 8 

        # 6. Costruzione Blocco Destro (Riassunto Consolidato)
        r_right = 1
        worksheet_stats.write(r_right, col_right, "", formati['header'])
        for i, year in enumerate(all_years):
            worksheet_stats.write(r_right, col_right + 1 + i, year, formati['header'])

        for m_idx, metric in enumerate(metrics_list):
            r_right_data = r_right + 1 + m_idx
            worksheet_stats.write(r_right_data, col_right, metric, formati['label'])
            for i, year in enumerate(all_years):
                if year in median_cells[metric]:
                    source_cell = median_cells[metric][year]
                    worksheet_stats.write_formula(r_right_data, col_right + 1 + i, f"={source_cell}", formati['data_num'])
                else:
                    worksheet_stats.write(r_right_data, col_right + 1 + i, "n.d.", formati['data_text'])

        # 7. Inserimento Grafici Dinamici (Separati)
        chart_main = workbook.add_chart({'type': 'column'})
        chart_sec = workbook.add_chart({'type': 'column'})
        
        metriche_isolate = ['Indice Rotazione Cap.Inv.', 'Gearing']
        has_sec_chart = False
        titolo_sec = ""

        for m_idx, metric in enumerate(metrics_list):
            r_idx = r_right + 1 + m_idx
            start_col = xlsxwriter.utility.xl_col_to_name(col_right + 1)
            end_col = xlsxwriter.utility.xl_col_to_name(col_right + len(all_years))
            
            serie = {
                'name':       f"='{sheet_stats}'!${xlsxwriter.utility.xl_col_to_name(col_right)}${r_idx + 1}",
                'categories': f"='{sheet_stats}'!${start_col}$2:${end_col}$2",
                'values':     f"='{sheet_stats}'!${start_col}${r_idx + 1}:${end_col}${r_idx + 1}",
                'data_labels': {'value': True, 'num_format': '#,##0.00'}
            }
            
            if metric in metriche_isolate:
                chart_sec.add_series(serie)
                has_sec_chart = True
                titolo_sec = f"Andamento Mediano {metric}"
            else:
                chart_main.add_series(serie)

        if metrics_list:
            chart_main.set_legend({'position': 'bottom'})
            chart_main.set_style(11)
            chart_main.set_y_axis({'name': y_axis_name, 'major_gridlines': {'visible': True}})
            chart_main.set_title({'name': chart_title})
            chart_main.set_size({'width': 600, 'height': 350})
            
            chart_row = r_right + len(metrics_list) + 3
            worksheet_stats.insert_chart(chart_row, col_right, chart_main)

            if has_sec_chart:
                chart_sec.set_legend({'position': 'bottom'})
                chart_sec.set_style(11)
                chart_sec.set_y_axis({'name': 'Valori', 'major_gridlines': {'visible': True}})
                chart_sec.set_title({'name': titolo_sec})
                chart_sec.set_size({'width': 600, 'height': 350})
                
                worksheet_stats.insert_chart(chart_row + 18, col_right, chart_sec)

        worksheet_stats.ignore_errors({'formula_differs': 'G1:Z500', 'number_stored_as_text': 'A1:Z500'})

    # ==========================================
    # LOGICA DI ESECUZIONE 
    # ==========================================
    df_raw = df_filtered.copy()
    output_buffer = io.BytesIO()
    
    writer = pd.ExcelWriter(output_buffer, engine='xlsxwriter')
    workbook = writer.book

    formati = {
        'header': workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
        'label': workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'left', 'valign': 'vcenter'}),
        'data_num': workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'align': 'right', 'valign': 'vcenter'}),
        'data_text': workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'}),
        'pct': workbook.add_format({'border': 1, 'num_format': '0.00%', 'align': 'right', 'valign': 'vcenter'})
    }

    costruisci_sezione_analisi(
        writer, workbook, df_raw, formati,
        keyword_ricerca=['struttura', 'gearing'], 
        sheet_data='5a. Dati Eq. Patrimoniale',
        sheet_stats='5b. Stat. Eq. Patrimoniale',
        chart_title='Andamento Mediano Indici',
        y_axis_name='Valori',
        rename_dict={
            'struttura 1° livello': 'Indice Struttura 1° Liv.',
            'struttura 2° livello': 'Indice Struttura 2° Liv.',
            'gearing': 'Gearing'
        }
    )

    writer.close()
    output_buffer.seek(0)
    
    return output_buffer


def elabora_capitolo_5(df_filtered):
    import io
    import pandas as pd
    import numpy as np
    import xlsxwriter
    from xlsxwriter.utility import xl_rowcol_to_cell
    import re

    def costruisci_sezione_analisi(writer, workbook, df_raw, formati, keyword_ricerca, sheet_data, sheet_stats, chart_title, y_axis_name, rename_dict):
        base_cols = [c for c in df_raw.columns if 'ragione' in str(c).lower() or 'bvd' in str(c).lower()]
        
        if isinstance(keyword_ricerca, list):
            metric_cols = [c for c in df_raw.columns if any(kw.lower() in str(c).lower() for kw in keyword_ricerca)]
        else:
            metric_cols = [c for c in df_raw.columns if keyword_ricerca.lower() in str(c).lower()]

        if not metric_cols:
            return

        df = df_raw[base_cols + metric_cols].copy()
        df = df.replace(['n.d.', 'n.a.', 'n.s.', ''], np.nan)
        num_rows = len(df)

        df.to_excel(writer, sheet_name=sheet_data, index=False, startrow=0)
        worksheet_data = writer.sheets[sheet_data]
        worksheet_stats = workbook.add_worksheet(sheet_stats)

        for i, col in enumerate(df.columns):
            worksheet_data.write(0, i, col, formati['header'])
            col_data = df[col].dropna()
            max_len = len(str(col)) + 2 if col_data.empty else max(col_data.astype(str).map(len).max(), len(str(col))) + 2
            worksheet_data.set_column(i, i, min(max_len, 45))
            for row in range(1, num_rows + 1):
                val = df.iat[row-1, i]
                if pd.isna(val):
                    worksheet_data.write(row, i, "", formati['data_text'])
                elif isinstance(val, (int, float)):
                    worksheet_data.write_number(row, i, val, formati['data_num'])
                else:
                    worksheet_data.write_string(row, i, str(val), formati['data_text'])

        numeric_cols_idx = [df.columns.get_loc(c) for c in metric_cols]
        metrics_dict = {}

        for col_idx in numeric_cols_idx:
            col_name = df.columns[col_idx]
            anno_match = re.search(r'\d{4}', col_name)
            if anno_match:
                anno = int(anno_match.group(0))
                metric_base = re.sub(r'\d{4}', '', col_name).replace('(*) %', '').replace('(*)', '').replace('migl EUR', '').replace(' ', ' ').strip()
                
                for k, v in rename_dict.items():
                    if k.lower() in metric_base.lower():
                        metric_base = v

                if metric_base not in metrics_dict: metrics_dict[metric_base] = {}
                metrics_dict[metric_base][anno] = col_idx

        all_years = sorted(list(set(anno for m in metrics_dict for anno in metrics_dict[m].keys())))
        metrics_list = list(metrics_dict.keys())

        col_left = 0
        col_mid = len(all_years) + 2
        col_right = col_mid + (len(all_years)*2 - 1) + 2

        worksheet_stats.set_column(col_left, col_left, 22)
        worksheet_stats.set_column(col_left + 1, col_left + len(all_years), 12)
        worksheet_stats.set_column(col_mid, col_mid, 22)
        worksheet_stats.set_column(col_mid + 1, col_mid + (len(all_years)*2), 10)
        worksheet_stats.set_column(col_right, col_right, 22)
        worksheet_stats.set_column(col_right + 1, col_right + len(all_years), 12)

        median_cells = {}
        current_row = 1
        stats_formulas = [('Media', 'AVERAGE'), ('Mediana', 'MEDIAN'), ('Asimmetria', 'SKEW'), ('Curtosi', 'KURT'), ('Deviazione standard', 'STDEV')]

        for metric in metrics_list:
            median_cells[metric] = {}

            worksheet_stats.write(current_row, col_left, "", formati['header'])
            for i, year in enumerate(all_years):
                worksheet_stats.write(current_row, col_left + 1 + i, year, formati['header'])

            for s_idx, (stat_name, func) in enumerate(stats_formulas):
                r = current_row + 1 + s_idx
                worksheet_stats.write(r, col_left, stat_name, formati['label'])
                for i, year in enumerate(all_years):
                    if year in metrics_dict[metric]:
                        data_col_idx = metrics_dict[metric][year]
                        col_letter = xlsxwriter.utility.xl_col_to_name(data_col_idx)
                        data_range = f"'{sheet_data}'!{col_letter}2:{col_letter}{num_rows+1}"
                        formula = f"={func}({data_range})"
                        worksheet_stats.write_formula(r, col_left + 1 + i, formula, formati['data_num'])
                        if stat_name == 'Mediana':
                            median_cells[metric][year] = xl_rowcol_to_cell(r, col_left + 1 + i)
                    else:
                        worksheet_stats.write(r, col_left + 1 + i, "n.d.", formati['data_text'])

            c_offset = 1
            for i, year in enumerate(all_years):
                if i == 0:
                    worksheet_stats.write(current_row, col_mid + c_offset, year, formati['header'])
                    worksheet_stats.write(current_row + 1, col_mid + c_offset, "V.m.", formati['header'])
                    c_offset += 1
                else:
                    worksheet_stats.merge_range(current_row, col_mid + c_offset, current_row, col_mid + c_offset + 1, year, formati['header'])
                    worksheet_stats.write(current_row + 1, col_mid + c_offset, "V.m.", formati['header'])
                    worksheet_stats.write(current_row + 1, col_mid + c_offset + 1, "Δ %", formati['header'])
                    c_offset += 2

            r_data = current_row + 2
            worksheet_stats.write(r_data, col_mid, metric, formati['label'])
            c_offset = 1
            prev_target_cell = None
            for i, year in enumerate(all_years):
                if year in median_cells[metric]:
                    source_cell = median_cells[metric][year]
                    curr_target_cell = xl_rowcol_to_cell(r_data, col_mid + c_offset)
                    worksheet_stats.write_formula(r_data, col_mid + c_offset, f"={source_cell}", formati['data_num'])

                    if i > 0:
                        if prev_target_cell:
                            delta_formula = f"=IFERROR(({curr_target_cell}-{prev_target_cell})/ABS({prev_target_cell}), 0)"
                            worksheet_stats.write_formula(r_data, col_mid + c_offset + 1, delta_formula, formati['pct'])
                        else:
                            worksheet_stats.write(r_data, col_mid + c_offset + 1, "n.d.", formati['data_text'])
                        prev_target_cell = curr_target_cell
                        c_offset += 2
                    else:
                        prev_target_cell = curr_target_cell
                        c_offset += 1
                else:
                    prev_target_cell = curr_target_cell
                    c_offset += 1

            current_row += 8 

        r_right = 1
        worksheet_stats.write(r_right, col_right, "", formati['header'])
        for i, year in enumerate(all_years):
            worksheet_stats.write(r_right, col_right + 1 + i, year, formati['header'])

        for m_idx, metric in enumerate(metrics_list):
            r_right_data = r_right + 1 + m_idx
            worksheet_stats.write(r_right_data, col_right, metric, formati['label'])
            for i, year in enumerate(all_years):
                if year in median_cells[metric]:
                    source_cell = median_cells[metric][year]
                    worksheet_stats.write_formula(r_right_data, col_right + 1 + i, f"={source_cell}", formati['data_num'])
                else:
                    worksheet_stats.write(r_right_data, col_right + 1 + i, "n.d.", formati['data_text'])

        chart_main = workbook.add_chart({'type': 'column'})
        chart_sec = workbook.add_chart({'type': 'column'})
        
        metriche_isolate = ['Indice Rotazione Cap.Inv.', 'Gearing']
        has_sec_chart = False
        titolo_sec = ""

        for m_idx, metric in enumerate(metrics_list):
            r_idx = r_right + 1 + m_idx
            start_col = xlsxwriter.utility.xl_col_to_name(col_right + 1)
            end_col = xlsxwriter.utility.xl_col_to_name(col_right + len(all_years))
            
            serie = {
                'name':       f"='{sheet_stats}'!${xlsxwriter.utility.xl_col_to_name(col_right)}${r_idx + 1}",
                'categories': f"='{sheet_stats}'!${start_col}$2:${end_col}$2",
                'values':     f"='{sheet_stats}'!${start_col}${r_idx + 1}:${end_col}${r_idx + 1}",
                'data_labels': {'value': True, 'num_format': '#,##0.00'}
            }
            
            if metric in metriche_isolate:
                chart_sec.add_series(serie)
                has_sec_chart = True
                titolo_sec = f"Andamento Mediano {metric}"
            else:
                chart_main.add_series(serie)

        if metrics_list:
            chart_main.set_legend({'position': 'bottom'})
            chart_main.set_style(11)
            chart_main.set_y_axis({'name': y_axis_name, 'major_gridlines': {'visible': True}})
            chart_main.set_title({'name': chart_title})
            chart_main.set_size({'width': 600, 'height': 350})
            
            chart_row = r_right + len(metrics_list) + 3
            worksheet_stats.insert_chart(chart_row, col_right, chart_main)

            if has_sec_chart:
                chart_sec.set_legend({'position': 'bottom'})
                chart_sec.set_style(11)
                chart_sec.set_y_axis({'name': 'Valori', 'major_gridlines': {'visible': True}})
                chart_sec.set_title({'name': titolo_sec})
                chart_sec.set_size({'width': 600, 'height': 350})
                
                worksheet_stats.insert_chart(chart_row + 18, col_right, chart_sec)

        worksheet_stats.ignore_errors({'formula_differs': 'G1:Z500', 'number_stored_as_text': 'A1:Z500'})

    df_raw = df_filtered.copy()
    output_buffer = io.BytesIO()
    
    writer = pd.ExcelWriter(output_buffer, engine='xlsxwriter')
    workbook = writer.book

    formati = {
        'header': workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
        'label': workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'left', 'valign': 'vcenter'}),
        'data_num': workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'align': 'right', 'valign': 'vcenter'}),
        'data_text': workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'}),
        'pct': workbook.add_format({'border': 1, 'num_format': '0.00%', 'align': 'right', 'valign': 'vcenter'})
    }

    costruisci_sezione_analisi(
        writer, workbook, df_raw, formati,
        keyword_ricerca=['current ratio', 'quick ratio', 'rotazione'], 
        sheet_data='6a. Dati Eq. Finanziario',
        sheet_stats='6b. Stat. Eq. Finanziario',
        chart_title='Andamento Mediano Indici',
        y_axis_name='Valori',
        rename_dict={
            'current ratio': 'Current Ratio',
            'quick ratio': 'Quick Ratio',
            'rotazione del capitale investito': 'Indice Rotazione Cap.Inv.'
        }
    )

    writer.close()
    output_buffer.seek(0)
    
    return output_buffer


def elabora_capitolo_6(df_filtered):
    import io
    import pandas as pd
    import numpy as np
    import xlsxwriter

    # 2. Caricamento dati (in memory)
    df_raw = df_filtered.copy()
    
    # 3. Definizione colonne
    def trova_col(keywords, exclude=None):
        if exclude is None: exclude = []
        for c in df_raw.columns:
            c_lower = str(c).lower()
            if all(k.lower() in c_lower for k in keywords) and not any(e.lower() in c_lower for e in exclude):
                return c
        return None

    col_regione = trova_col(['nuts', '2'])
    if not col_regione: col_regione = trova_col(['nuts'])
    
    col_macro = trova_col(['nuts', '1'])
    if not col_macro:
        if col_regione:
            df_raw['Macro_Fallback'] = df_raw[col_regione].astype(str).str[:3] + " - Macroregione"
            col_macro = 'Macro_Fallback'
        else:
             df_raw['Macro_Fallback'] = "N.D. - Macroregione"
             col_macro = 'Macro_Fallback'

    cols_dict = {
        'Ragione Sociale': trova_col(['ragione']),
        'Macro-Regione': col_macro, 
        'Regione': col_regione,
        'M. Profitto 2024': trova_col(['marg', 'profitto', '2024']),
        'M. EBITDA 2024': trova_col(['marg', 'ebitda', '2024']),
        'M. EBIT 2024': trova_col(['marg', 'ebit', '2024'], exclude=['ebitda']),
        'Rotazione C.Inv. 2024': trova_col(['rotazione', '2024']),
        'Quick Ratio 2024': trova_col(['quick', '2024']),
        'Current Ratio 2024': trova_col(['current', '2024']),
        'Indice 1° Liv. 2024': trova_col(['struttura 1', '2024']),
        'Indice 2° Liv. 2024': trova_col(['struttura 2', '2024']),
        'Gearing 2024': trova_col(['gearing', '2024'])
    }

    if cols_dict['Regione'] is None:
        df_raw['Regione_Fallback'] = "N.D."
        cols_dict['Regione'] = 'Regione_Fallback'

    df = df_raw[list(cols_dict.values())].copy()
    df.columns = list(cols_dict.keys())
    
    metriche = list(cols_dict.keys())[3:] 
    for m in metriche:
        df[m] = pd.to_numeric(df[m], errors='coerce')

    # 4. Creazione Colonne Vuote per Spaziature e Formule
    df.insert(1, 'Società ID', range(1, len(df) + 1))
    
    df[' '] = ''    
    df['  '] = ''   
    df['   '] = ''  
    df['    '] = '' 

    df['Benchmark Economico'] = ''
    df['Benchmark Finanziario'] = ''
    df['Benchmark Patrimoniale'] = ''
    df['Benchmark Totale'] = ''
    df['Rating Combinato'] = ''

    # 5. Ordinamento colonne
    col_finali = [
        'Ragione Sociale', 'Società ID',
        ' ', 
        'M. Profitto 2024', 'M. EBITDA 2024', 'M. EBIT 2024',
        '  ', 
        'Rotazione C.Inv. 2024', 'Quick Ratio 2024', 'Current Ratio 2024',
        '   ', 
        'Indice 1° Liv. 2024', 'Indice 2° Liv. 2024', 'Gearing 2024',
        '    ', 
        'Benchmark Economico', 'Benchmark Finanziario', 'Benchmark Patrimoniale', 'Benchmark Totale',
        'Regione', 'Rating Combinato'
    ]
    df_out = df[col_finali]

    start_col_tbl = len(df_out.columns) + 1 
    col_T1 = xlsxwriter.utility.xl_col_to_name(start_col_tbl + 2) 
    col_T2 = xlsxwriter.utility.xl_col_to_name(start_col_tbl + 3) 

    output_buffer = io.BytesIO()
    
    # 6. Scrittura su Excel (Foglio 7)
    writer = pd.ExcelWriter(output_buffer, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet('7. Rating e Benchmark 2024')
    
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    fmt_header_metric = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    fmt_data = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
    fmt_num = workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'align': 'right', 'valign': 'vcenter'})
    fmt_center = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_id = workbook.add_format({'border': 1, 'num_format': '0', 'align': 'center', 'valign': 'vcenter'}) 
    fmt_space = workbook.add_format({'bg_color': '#EAEAEA'}) 

    for col_num, col_name in enumerate(df_out.columns):
        if col_name.strip() == '':
            worksheet.write(0, col_num, "", fmt_space)
        else:
            formato = fmt_header_metric if '2024' in col_name else fmt_header
            worksheet.write(0, col_num, col_name, formato)
        
    for row_num, row_data in enumerate(df_out.values):
        row_ex = row_num + 2 
        for col_num, val in enumerate(row_data):
            col_name = df_out.columns[col_num]
            
            if col_name.strip() == '':
                worksheet.write(row_ex - 1, col_num, "", fmt_space)
            
            elif col_name == 'Benchmark Economico':
                cond_D = f"IF(D{row_ex}>=${col_T2}$4,3,IF(D{row_ex}>=${col_T1}$4,2,1))"
                cond_E = f"IF(E{row_ex}>=${col_T2}$5,3,IF(E{row_ex}>=${col_T1}$5,2,1))"
                cond_F = f"IF(F{row_ex}>=${col_T2}$6,3,IF(F{row_ex}>=${col_T1}$6,2,1))"
                formula = f'=IF(({cond_D}+{cond_E}+{cond_F})>=8,"A",IF(({cond_D}+{cond_E}+{cond_F})>=5,"B","C"))'
                worksheet.write_formula(row_ex - 1, col_num, formula, fmt_center)
                
            elif col_name == 'Benchmark Finanziario':
                cond_H = f"IF(H{row_ex}<=${col_T1}$7,3,IF(H{row_ex}<=${col_T2}$7,2,1))" 
                cond_I = f"IF(I{row_ex}>=${col_T2}$8,3,IF(I{row_ex}>=${col_T1}$8,2,1))"
                cond_J = f"IF(J{row_ex}>=${col_T2}$9,3,IF(J{row_ex}>=${col_T1}$9,2,1))"
                formula = f'=IF(({cond_H}+{cond_I}+{cond_J})>=8,"A",IF(({cond_H}+{cond_I}+{cond_J})>=5,"B","C"))'
                worksheet.write_formula(row_ex - 1, col_num, formula, fmt_center)
                
            elif col_name == 'Benchmark Patrimoniale':
                cond_L = f"IF(L{row_ex}>=${col_T2}$10,3,IF(L{row_ex}>=${col_T1}$10,2,1))"
                cond_M = f"IF(M{row_ex}>=${col_T2}$11,3,IF(M{row_ex}>=${col_T1}$11,2,1))"
                cond_N = f"IF(N{row_ex}<=${col_T1}$12,3,IF(N{row_ex}<=${col_T2}$12,2,1))"
                formula = f'=IF(({cond_L}+{cond_M}+{cond_N})>=8,"A",IF(({cond_L}+{cond_M}+{cond_N})>=5,"B","C"))'
                worksheet.write_formula(row_ex - 1, col_num, formula, fmt_center)
                
            elif col_name == 'Benchmark Totale':
                cond_P = f'IF(P{row_ex}="A",3,IF(P{row_ex}="B",2,1))'
                cond_Q = f'IF(Q{row_ex}="A",3,IF(Q{row_ex}="B",2,1))'
                cond_R = f'IF(R{row_ex}="A",3,IF(R{row_ex}="B",2,1))'
                formula = f'=IF(({cond_P}+{cond_Q}+{cond_R})>=8,"A",IF(({cond_P}+{cond_Q}+{cond_R})>=5,"B","C"))'
                worksheet.write_formula(row_ex - 1, col_num, formula, fmt_center)
                
            elif col_name == 'Rating Combinato':
                formula = f'=P{row_ex}&Q{row_ex}&R{row_ex}'
                worksheet.write_formula(row_ex - 1, col_num, formula, fmt_center)

            elif pd.isna(val):
                worksheet.write(row_ex - 1, col_num, "n.d.", fmt_data)
            elif col_name == 'Società ID': 
                worksheet.write_number(row_ex - 1, col_num, val, fmt_id)
            elif isinstance(val, (int, float)):
                worksheet.write_number(row_ex - 1, col_num, val, fmt_num)
            else:
                worksheet.write(row_ex - 1, col_num, str(val), fmt_data)

    for col_num, col_name in enumerate(df_out.columns):
        if col_name.strip() == '':
            worksheet.set_column(col_num, col_num, 2) 
            
    worksheet.set_column('A:A', 35)
    worksheet.set_column('B:B', 10)
    worksheet.set_column('D:F', 14) 
    worksheet.set_column('H:J', 14) 
    worksheet.set_column('L:N', 14) 
    worksheet.set_column('P:S', 20) 
    worksheet.set_column('T:T', 35) 
    worksheet.set_column('U:U', 18) 
    worksheet.set_row(0, 45) 

    worksheet.write(1, start_col_tbl, "SOGLIE CALCOLATE 2024", fmt_header)
    worksheet.write(2, start_col_tbl, "Metrica", fmt_header)
    intestazioni_tbl = ["MIN", "Soglia 1° Terzile", "Soglia 2° Terzile", "MAX"]
    for i, h in enumerate(intestazioni_tbl):
        worksheet.write(2, start_col_tbl + 1 + i, h, fmt_header)

    num_rows = len(df_out)
    for i, m in enumerate(metriche):
        riga = 3 + i
        worksheet.write(riga, start_col_tbl, m, fmt_header_metric)
        col_idx = df_out.columns.get_loc(m)
        col_letter = xlsxwriter.utility.xl_col_to_name(col_idx)
        data_range = f"{col_letter}2:{col_letter}{num_rows + 1}"
        worksheet.write_formula(riga, start_col_tbl + 1, f"=MIN({data_range})", fmt_num)
        worksheet.write_formula(riga, start_col_tbl + 2, f"=PERCENTILE({data_range}, 1/3)", fmt_num)
        worksheet.write_formula(riga, start_col_tbl + 3, f"=PERCENTILE({data_range}, 2/3)", fmt_num)
        worksheet.write_formula(riga, start_col_tbl + 4, f"=MAX({data_range})", fmt_num)
        
    worksheet.set_column(start_col_tbl, start_col_tbl, 20)
    worksheet.set_column(start_col_tbl + 1, start_col_tbl + 4, 18)

    # ---------------------------------------------------------
    # AGGIUNTA FOGLI PIVOT E RIPARTIZIONE MACRO/NUTS2
    # ---------------------------------------------------------
    fmt_pct = workbook.add_format({'border': 1, 'num_format': '0.00%', 'align': 'right', 'valign': 'vcenter'})
    fmt_num_int = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter'})
    
    fmt_subtotal = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
    fmt_subtotal_num = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter'})
    fmt_subtotal_pct = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'num_format': '0.00%', 'align': 'right', 'valign': 'vcenter'})
    fmt_grand_tot = workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
    fmt_grand_tot_num = workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter'})
    fmt_grand_tot_pct = workbook.add_format({'bold': True, 'bg_color': '#002060', 'font_color': 'white', 'border': 1, 'num_format': '0.00%', 'align': 'right', 'valign': 'vcenter'})

    bench_cols = {
        'Benchmark Economico': 'P',
        'Benchmark Finanziario': 'Q',
        'Benchmark Patrimoniale': 'R',
        'Benchmark Totale': 'S'
    }
    main_sheet = "'7. Rating e Benchmark 2024'"
    range_reg = f"{main_sheet}!$T$2:$T${num_rows+1}"

    # --- Foglio 8: Pivot Analisi Benchmark ---
    worksheet_pivot = workbook.add_worksheet('8. Pivot Analisi Benchmark')
    worksheet_pivot.set_column('A:A', 35)
    worksheet_pivot.set_column('B:C', 15)

    row_p = 1
    for b_name, b_col in bench_cols.items():
        worksheet_pivot.write(row_p, 0, f"Conteggio di {b_name}", fmt_header)
        worksheet_pivot.write(row_p, 1, "V.A.", fmt_header)
        worksheet_pivot.write(row_p, 2, "%", fmt_header)
        row_p += 1

        range_bench = f"{main_sheet}!${b_col}$2:${b_col}${num_rows+1}"
        row_start = row_p

        for lettera in ['A', 'B', 'C']:
            worksheet_pivot.write(row_p, 0, lettera, fmt_center)
            worksheet_pivot.write_formula(row_p, 1, f'=COUNTIF({range_bench}, "{lettera}")', fmt_num_int)
            row_p += 1

        row_tot = row_p
        worksheet_pivot.write(row_tot, 0, "Totale complessivo", fmt_header)
        worksheet_pivot.write_formula(row_tot, 1, f'=SUM(B{row_start+1}:B{row_tot})', fmt_num_int)
        worksheet_pivot.write(row_tot, 2, 1.0, fmt_pct)

        for i in range(3):
            worksheet_pivot.write_formula(row_start + i, 2, f'=IF($B${row_tot+1}>0, B{row_start+i+1}/$B${row_tot+1}, 0)', fmt_pct)

        row_p += 3

    # --- Foglio 9: Rip.Terr. Benchmark ---
    worksheet_terr = workbook.add_worksheet('9. Rip.Terr. Benchmark')
    worksheet_terr.set_column('A:A', 40)
    worksheet_terr.set_column('B:I', 15)

    mapping_df = df[['Macro-Regione', 'Regione']].dropna()
    mapping_df = mapping_df[(mapping_df['Regione'].astype(str).str.strip() != '') & (mapping_df['Regione'].astype(str).str.lower() != 'nan')]
    
    mapping = {}
    for macro, group in mapping_df.groupby('Macro-Regione'):
        macro_str = str(macro).strip()
        if macro_str and macro_str.lower() != 'nan':
            mapping[macro_str] = group['Regione'].unique().tolist()

    sorted_macro = sorted(mapping.keys())

    row_t = 1
    for b_name, b_col in bench_cols.items():
        range_bench = f"{main_sheet}!${b_col}$2:${b_col}${num_rows+1}"

        worksheet_terr.merge_range(row_t, 0, row_t, 8, f"{b_name} - Ripartizione Territoriale", fmt_header_metric)
        row_t += 1

        headers = ["Regioni", "Imprese V.A.", "Imprese %", "A (V.A.)", "A (%)", "B (V.A.)", "B (%)", "C (V.A.)", "C (%)"]
        for i, h in enumerate(headers):
            worksheet_terr.write(row_t, i, h, fmt_header)
        row_t += 1

        offset_righe = sum(len(mapping[m]) + 2 for m in sorted_macro) - 1 if sorted_macro else 0
        riga_italia = row_t + offset_righe
        total_cell = f"$B${riga_italia+1}"
        
        grand_total_row_refs = []

        for macro in sorted_macro:
            regioni_list = mapping[macro]
            start_macro_row = row_t

            for reg in sorted(regioni_list):
                worksheet_terr.write(row_t, 0, str(reg), fmt_data)
                worksheet_terr.write_formula(row_t, 1, f'=COUNTIF({range_reg}, "{reg}")', fmt_num_int)
                worksheet_terr.write_formula(row_t, 2, f'=IF({total_cell}>0, B{row_t+1}/{total_cell}, 0)', fmt_pct)
                worksheet_terr.write_formula(row_t, 3, f'=COUNTIFS({range_reg}, "{reg}", {range_bench}, "A")', fmt_num_int)
                worksheet_terr.write_formula(row_t, 4, f'=IF({total_cell}>0, D{row_t+1}/{total_cell}, 0)', fmt_pct)
                worksheet_terr.write_formula(row_t, 5, f'=COUNTIFS({range_reg}, "{reg}", {range_bench}, "B")', fmt_num_int)
                worksheet_terr.write_formula(row_t, 6, f'=IF({total_cell}>0, F{row_t+1}/{total_cell}, 0)', fmt_pct)
                worksheet_terr.write_formula(row_t, 7, f'=COUNTIFS({range_reg}, "{reg}", {range_bench}, "C")', fmt_num_int)
                worksheet_terr.write_formula(row_t, 8, f'=IF({total_cell}>0, H{row_t+1}/{total_cell}, 0)', fmt_pct)
                row_t += 1

            worksheet_terr.write(row_t, 0, str(macro), fmt_subtotal)
            worksheet_terr.write_formula(row_t, 1, f'=SUM(B{start_macro_row+1}:B{row_t})', fmt_subtotal_num)
            worksheet_terr.write_formula(row_t, 2, f'=IF({total_cell}>0, B{row_t+1}/{total_cell}, 0)', fmt_subtotal_pct)
            worksheet_terr.write_formula(row_t, 3, f'=SUM(D{start_macro_row+1}:D{row_t})', fmt_subtotal_num)
            worksheet_terr.write_formula(row_t, 4, f'=IF({total_cell}>0, D{row_t+1}/{total_cell}, 0)', fmt_subtotal_pct)
            worksheet_terr.write_formula(row_t, 5, f'=SUM(F{start_macro_row+1}:F{row_t})', fmt_subtotal_num)
            worksheet_terr.write_formula(row_t, 6, f'=IF({total_cell}>0, F{row_t+1}/{total_cell}, 0)', fmt_subtotal_pct)
            worksheet_terr.write_formula(row_t, 7, f'=SUM(H{start_macro_row+1}:H{row_t})', fmt_subtotal_num)
            worksheet_terr.write_formula(row_t, 8, f'=IF({total_cell}>0, H{row_t+1}/{total_cell}, 0)', fmt_subtotal_pct)

            grand_total_row_refs.append(row_t + 1)
            row_t += 2 

        row_tot = row_t - 1 

        worksheet_terr.write(row_tot, 0, "Italia", fmt_grand_tot)
        if grand_total_row_refs:
            sum_b = "+".join([f"B{r}" for r in grand_total_row_refs])
            sum_d = "+".join([f"D{r}" for r in grand_total_row_refs])
            sum_f = "+".join([f"F{r}" for r in grand_total_row_refs])
            sum_h = "+".join([f"H{r}" for r in grand_total_row_refs])
            worksheet_terr.write_formula(row_tot, 1, f'={sum_b}', fmt_grand_tot_num)
            worksheet_terr.write_formula(row_tot, 3, f'={sum_d}', fmt_grand_tot_num)
            worksheet_terr.write_formula(row_tot, 5, f'={sum_f}', fmt_grand_tot_num)
            worksheet_terr.write_formula(row_tot, 7, f'={sum_h}', fmt_grand_tot_num)
        else:
             worksheet_terr.write(row_tot, 1, 0, fmt_grand_tot_num)
             worksheet_terr.write(row_tot, 3, 0, fmt_grand_tot_num)
             worksheet_terr.write(row_tot, 5, 0, fmt_grand_tot_num)
             worksheet_terr.write(row_tot, 7, 0, fmt_grand_tot_num)

        worksheet_terr.write_formula(row_tot, 2, f'=IF({total_cell}>0, B{row_tot+1}/{total_cell}, 0)', fmt_grand_tot_pct)
        worksheet_terr.write_formula(row_tot, 4, f'=IF({total_cell}>0, D{row_tot+1}/{total_cell}, 0)', fmt_grand_tot_pct)
        worksheet_terr.write_formula(row_tot, 6, f'=IF({total_cell}>0, F{row_tot+1}/{total_cell}, 0)', fmt_grand_tot_pct)
        worksheet_terr.write_formula(row_tot, 8, f'=IF({total_cell}>0, H{row_tot+1}/{total_cell}, 0)', fmt_grand_tot_pct)

        row_t += 4

    worksheet_terr.ignore_errors({'formula_differs': 'A1:Z1000'})

    writer.close()
    output_buffer.seek(0)
    
    return output_buffer


def elabora_capitolo_7(df_filtered):
    import io
    import pandas as pd
    import numpy as np
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    import zipfile

    # --- FUNZIONI DI SUPPORTO ---
    def calcola_punteggi_diretto(val, t1, t2):
        if pd.isna(val): return 1
        if val >= t2: return 3
        elif val >= t1: return 2
        else: return 1

    def calcola_punteggi_inverso(val, t1, t2):
        if pd.isna(val): return 1
        if val <= t1: return 3
        elif val <= t2: return 2
        else: return 1

    def assegna_lettera(punti, soglia_A=8, soglia_B=5):
        if pd.isna(punti): return 'C'
        if punti >= soglia_A: return 'A'
        elif punti >= soglia_B: return 'B'
        else: return 'C'

    def punti_da_lettera(lettera):
        if lettera == 'A': return 3
        elif lettera == 'B': return 2
        else: return 1

    df_raw = df_filtered.copy()

    def trova_col(keywords, exclude=None):
        if exclude is None: exclude = []
        for c in df_raw.columns:
            c_lower = str(c).lower()
            if all(k.lower() in c_lower for k in keywords) and not any(e.lower() in c_lower for e in exclude):
                return c
        return None

    col_regione = trova_col(['nuts', '2'])
    if not col_regione: col_regione = trova_col(['nuts'])
    if not col_regione:
        df_raw['Regione_Fallback'] = "N.D."
        col_regione = 'Regione_Fallback'

    col_societa = 'Società ID'
    df_raw.insert(1, col_societa, range(1, len(df_raw) + 1))
    
    col_ragione_sociale = trova_col(['ragione'])

    # Ricreiamo le metriche usate per il ranking
    metriche_utili = {
        'M. Profitto 2024': trova_col(['marg', 'profitto', '2024']),
        'M. EBITDA 2024': trova_col(['marg', 'ebitda', '2024']),
        'M. EBIT 2024': trova_col(['marg', 'ebit', '2024'], exclude=['ebitda']),
        'Rotazione C.Inv. 2024': trova_col(['rotazione', '2024']),
        'Quick Ratio 2024': trova_col(['quick', '2024']),
        'Current Ratio 2024': trova_col(['current', '2024']),
        'Indice 1° Liv. 2024': trova_col(['struttura 1', '2024']),
        'Indice 2° Liv. 2024': trova_col(['struttura 2', '2024']),
        'Gearing 2024': trova_col(['gearing', '2024'])
    }

    df = pd.DataFrame()
    df[col_ragione_sociale] = df_raw[col_ragione_sociale]
    df[col_societa] = df_raw[col_societa]
    df[col_regione] = df_raw[col_regione]
    
    for nome_bello, col_brutta in metriche_utili.items():
        if col_brutta:
            df[nome_bello] = pd.to_numeric(df_raw[col_brutta], errors='coerce')
        else:
            df[nome_bello] = np.nan

    col_lettere = ['Benchmark Economico', 'Benchmark Finanziario', 'Benchmark Patrimoniale', 'Benchmark Totale', 'Rating Combinato']

    categorie_kpi = {
        'Equilibrio_Economico': {
            'Prof_Mg': ('M. Profitto 2024', False),
            'EBITDA_Mg': ('M. EBITDA 2024', False),
            'EBIT_Mg': ('M. EBIT 2024', False)
        },
        'Equilibrio_Finanziario': {
            'Rotazione_Cap': ('Rotazione C.Inv. 2024', True),
            'Quick_Rat': ('Quick Ratio 2024', False),
            'Current_Rat': ('Current Ratio 2024', False)
        },
        'Equilibrio_Patrimoniale': {
            'IndStrut1': ('Indice 1° Liv. 2024', False),
            'IndStrut2': ('Indice 2° Liv. 2024', False),
            'Gearing': ('Gearing 2024', True) 
        }
    }

    tutti_kpi_cols = []
    for kpi_dict in categorie_kpi.values():
        tutti_kpi_cols.extend([v[0] for v in kpi_dict.values()])

    metriche_inverse = ['Rotazione C.Inv. 2024', 'Gearing 2024']

    for m in tutti_kpi_cols:
        if m in df.columns:
            t1 = df[m].quantile(1/3)
            t2 = df[m].quantile(2/3)
            if m in metriche_inverse:
                df[f'Pts_{m}'] = df[m].apply(lambda x: calcola_punteggi_inverso(x, t1, t2))
            else:
                df[f'Pts_{m}'] = df[m].apply(lambda x: calcola_punteggi_diretto(x, t1, t2))

    if all(f'Pts_{x}' in df.columns for x in ['M. Profitto 2024', 'M. EBITDA 2024', 'M. EBIT 2024']):
        df['Sum_Eco'] = df['Pts_M. Profitto 2024'] + df['Pts_M. EBITDA 2024'] + df['Pts_M. EBIT 2024']
        df['Benchmark Economico'] = df['Sum_Eco'].apply(lambda x: assegna_lettera(x, 8, 5))
    else: df['Benchmark Economico'] = 'C'

    if all(f'Pts_{x}' in df.columns for x in ['Rotazione C.Inv. 2024', 'Quick Ratio 2024', 'Current Ratio 2024']):
        df['Sum_Fin'] = df['Pts_Rotazione C.Inv. 2024'] + df['Pts_Quick Ratio 2024'] + df['Pts_Current Ratio 2024']
        df['Benchmark Finanziario'] = df['Sum_Fin'].apply(lambda x: assegna_lettera(x, 8, 5))
    else: df['Benchmark Finanziario'] = 'C'

    if all(f'Pts_{x}' in df.columns for x in ['Indice 1° Liv. 2024', 'Indice 2° Liv. 2024', 'Gearing 2024']):
        df['Sum_Pat'] = df['Pts_Indice 1° Liv. 2024'] + df['Pts_Indice 2° Liv. 2024'] + df['Pts_Gearing 2024']
        df['Benchmark Patrimoniale'] = df['Sum_Pat'].apply(lambda x: assegna_lettera(x, 8, 5))
    else: df['Benchmark Patrimoniale'] = 'C'

    df['Sum_Lettere'] = df['Benchmark Economico'].apply(punti_da_lettera) + \
                        df['Benchmark Finanziario'].apply(punti_da_lettera) + \
                        df['Benchmark Patrimoniale'].apply(punti_da_lettera)

    df['Benchmark Totale'] = df['Sum_Lettere'].apply(lambda x: assegna_lettera(x, 8, 5))
    df['Rating Combinato'] = df['Benchmark Economico'] + df['Benchmark Finanziario'] + df['Benchmark Patrimoniale']

    df_master = df.dropna(subset=[col_ragione_sociale]).copy()
    colonne_da_tenere = [col_ragione_sociale, col_societa, col_regione] + [c for c in col_lettere if c in df.columns] + tutti_kpi_cols

    def calcola_rank(dataframe, col_kpi, base_name, lower_is_better):
        asc_order = True if lower_is_better else False
        dataframe[f'RANK_{base_name}_NAZ'] = dataframe[col_kpi].rank(ascending=asc_order, method='min')
        dataframe[f'RANK_{base_name}_REG'] = dataframe.groupby(col_regione, observed=False)[col_kpi].rank(ascending=asc_order, method='min')

    # Creazione del buffer ZIP per ospitare i 3 file Excel
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for categoria, kpi_dict in categorie_kpi.items():
            file_name = f'7_Ranking_Aziendale_{categoria}.xlsx'
            df_cat = df_master.copy()
            
            for kpi_name, (kpi_col, is_lower_better) in kpi_dict.items():
                if kpi_col in df_cat.columns:
                    calcola_rank(df_cat, kpi_col, kpi_name, is_lower_better)
            
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                if 'Rating Combinato' in df_cat.columns:
                    colonne_rank = [c for c in df_cat.columns if '_NAZ' in c]
                    col_rank_spareggio = colonne_rank[0] if colonne_rank else col_societa
                    
                    df_rating = df_cat.sort_values(
                        by=['Rating Combinato', 'Benchmark Totale', col_rank_spareggio], 
                        ascending=[True, True, True] 
                    )
                    cols = [col_ragione_sociale] + [c for c in df_rating.columns if c in colonne_da_tenere + colonne_rank and c != col_ragione_sociale]
                    df_rating[cols].to_excel(writer, sheet_name='TOP_RATING_ABC', index=False)
                
                for kpi_name, (kpi_col, _) in kpi_dict.items():
                    if kpi_col in df_cat.columns:
                        df_naz = df_cat.sort_values(by=f'RANK_{kpi_name}_NAZ').dropna(subset=[kpi_col])
                        cols_naz = [col_ragione_sociale] + [c for c in df_naz.columns if c in colonne_da_tenere + colonne_rank and c != col_ragione_sociale]
                        df_naz[cols_naz].to_excel(writer, sheet_name=f'{kpi_name}_Nazionale', index=False)
                        
                        df_reg = df_cat.sort_values(by=[col_regione, f'RANK_{kpi_name}_REG']).dropna(subset=[kpi_col])
                        cols_reg = [col_ragione_sociale] + [c for c in df_reg.columns if c in colonne_da_tenere + colonne_rank and c != col_ragione_sociale]
                        df_reg[cols_reg].to_excel(writer, sheet_name=f'{kpi_name}_Regionale', index=False)

            excel_buffer.seek(0)
            wb = openpyxl.load_workbook(excel_buffer)
            header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            white_font = Font(color="FFFFFF", bold=True)
            center_align = Alignment(horizontal='center')
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            for ws in wb.worksheets:
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = white_font
                    cell.alignment = center_align
                    cell.border = thin_border

                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    
                    if col[0].value and ("RANK" in str(col[0].value) or "Rating" in str(col[0].value) or "Benchmark" in str(col[0].value)):
                        for cell in col:
                            cell.alignment = center_align
                    
                    for cell in col:
                        try:
                            if cell.value and len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except: pass
                    ws.column_dimensions[column].width = min(max_length + 3, 50)

            final_excel_buffer = io.BytesIO()
            wb.save(final_excel_buffer)
            final_excel_buffer.seek(0)
            
            zip_file.writestr(file_name, final_excel_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer


# ==========================================
# 2. INTERFACCIA WEB (Il "Capitolo 0" + Filtri)
# ==========================================

st.info("💡 **ATTENZIONE:** Il file caricato deve essere un export diretto da **ORBIS** in formato `.xlsx`, generato utilizzando gli specifici filtri e il formato di esportazione **LISTA UNIVERSAL**.")

uploaded_file = st.file_uploader("Trascina qui l'export grezzo di ORBIS (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Verifica formato, lettura e pulizia dati in corso..."):
            
        # --- 1. VERIFICA PRESENZA DEL FOGLIO CORRETTO ---
        try:
            xls = pd.ExcelFile(uploaded_file, engine='calamine')
        except Exception as e:
            st.error("❌ ERRORE: Impossibile leggere il file Excel.")
            st.warning(f"🛠️ DETTAGLIO TECNICO: {e}")
            st.stop()
            
        fogli_disponibili = xls.sheet_names
        target_sheet = None
        if "Risultati" in fogli_disponibili:
            target_sheet = "Risultati"
        elif "Results" in fogli_disponibili:
            target_sheet = "Results"
            
        if target_sheet is None:
            st.error("❌ ERRORE: Foglio 'Risultati' o 'Results' non trovato. ESPORTARE DA ORBIS CON I FILTRI E LA LISTA UNIVERSAL.")
            st.stop()

        # Lettura del foglio corretto
        df_orbis = pd.read_excel(xls, sheet_name=target_sheet)

        # --- 2. VERIFICA STRUTTURA DELLE COLONNE ---
        # La tua "Lista Universal" rigorosa
        colonne_attese = [
            'Ragione socialeCaratteri latini', 'Numero BvD ID', 'Forma giuridica nazionale', 
            'NUTS1', 'NUTS2', 'NUTS3', 'Numero dipendenti 2024', 
            'Totale valore della produzione migl EUR 2024', 'Totale valore della produzione migl EUR 2023', 'Totale valore della produzione migl EUR 2022', 'Totale valore della produzione migl EUR 2021', 
            'Totale Attivo migl EUR 2024', 'Totale Attivo migl EUR 2023', 'Totale Attivo migl EUR 2022', 'Totale Attivo migl EUR 2021', 
            'Margine di Profitto (*) % 2024', 'Margine di Profitto (*) % 2023', 'Margine di Profitto (*) % 2022', 'Margine di Profitto (*) % 2021', 
            'Margine EBITDA (*) % 2024', 'Margine EBITDA (*) % 2023', 'Margine EBITDA (*) % 2022', 'Margine EBITDA (*) % 2021', 
            'Margine EBIT (*) % 2024', 'Margine EBIT (*) % 2023', 'Margine EBIT (*) % 2022', 'Margine EBIT (*) % 2021', 
            'Indice di Struttura 1° livello (*) 2024', 'Indice di Struttura 1° livello (*) 2023', 'Indice di Struttura 1° livello (*) 2022', 'Indice di Struttura 1° livello (*) 2021', 
            'Indice di Struttura 2° livello (*) 2024', 'Indice di Struttura 2° livello (*) 2023', 'Indice di Struttura 2° livello (*) 2022', 'Indice di Struttura 2° livello (*) 2021', 
            'Gearing (*) % 2024', 'Gearing (*) % 2023', 'Gearing (*) % 2022', 'Gearing (*) % 2021', 
            'Current Ratio (*) 2024', 'Current Ratio (*) 2023', 'Current Ratio (*) 2022', 'Current Ratio (*) 2021', 
            'Quick Ratio (*) 2024', 'Quick Ratio (*) 2023', 'Quick Ratio (*) 2022', 'Quick Ratio (*) 2021', 
            'Indice di Rotazione del Capitale Investito (*) 2024', 'Indice di Rotazione del Capitale Investito (*) 2023', 'Indice di Rotazione del Capitale Investito (*) 2022', 'Indice di Rotazione del Capitale Investito (*) 2021'
        ]
            
        # Pulisce gli spazi laterali dalle colonne per un controllo accurato
        colonne_file = [str(c).strip() for c in df_orbis.columns]
        colonne_mancanti = [col for col in colonne_attese if col not in colonne_file]
            
        if len(colonne_mancanti) > 0:
            st.error("❌ ERRORE: Colonne non valide o mancanti. ESPORTARE DA ORBIS CON I FILTRI E LA LISTA UNIVERSAL.")
            # Opzionale: un menu a tendina per far vedere esattamente cosa manca (utile per il debug)
            with st.expander("Mostra i dettagli dell'errore (Colonne Mancanti)"):
                st.write(colonne_mancanti)
            st.stop()  # IL PROGRAMMA SI FERMA QUI, I PULSANTI NON VERRANNO GENERATI

        # --- 3. PULIZIA DATI E FILTRAGGIO VALORI (n.d. e Rotazione) ---
        righe_iniziali = len(df_orbis)

        col_att_24 = 'Totale Attivo migl EUR 2024'
        col_ric_24 = 'Totale valore della produzione migl EUR 2024'
        col_rot_24 = 'Indice di Rotazione del Capitale Investito (*) 2024'

        # Forza la conversione a numero: tutto ciò che è testo (come "n.d.") diventa vuoto (NaN)
        for col in [col_att_24, col_ric_24, col_rot_24]:
            df_orbis[col] = pd.to_numeric(df_orbis[col], errors='coerce')

        # Elimina le righe che non hanno dati al 2024 su Ricavi, Attivo o Rotazione
        df_orbis = df_orbis.dropna(subset=[col_att_24, col_ric_24, col_rot_24])
            
        # Mantieni solo aziende con Rotazione strettamente maggiore di 0
        df_orbis = df_orbis[df_orbis[col_rot_24] > 0]
            
        # --- 4. RENUMERAZIONE PROGRESSIVA DELLA PRIMA COLONNA ---
        # Resetta l'indice del dataframe eliminando i "buchi"
        df_orbis = df_orbis.reset_index(drop=True)
            
        # La primissima colonna di ORBIS (spesso chiamata "Unnamed: 0") contiene la numerazione
        prima_colonna = df_orbis.columns[0] 
        if "Ragione" not in str(prima_colonna): # Controllo di sicurezza
            # Rigenera la lista con i nuovi posti scalati (es. "1.", "2.", "3.")
            df_orbis[prima_colonna] = [f"{i}." for i in range(1, len(df_orbis) + 1)]

        righe_finali = len(df_orbis)
        righe_scartate = righe_iniziali - righe_finali

    # Mostra i risultati del filtro a schermo
    if righe_scartate > 0:
        st.warning(f"🧹 **Pulizia automatica:** Rimosse **{righe_scartate}** aziende perché mancavano dati al 2024 o la rotazione era ≤ 0. L'ordine è stato ricalcolato. (Aziende valide: **{righe_finali}**)")
    else:
        st.success(f"✅ **Dati perfetti!** Tutte le {righe_finali} aziende hanno passato i controlli e sono pronte all'uso.")


    # ==========================================
    # 3. GESTIONE DELLE TABS (I Pulsanti)
    # ==========================================
    
    # Crea le schede per i vari capitoli + Tab per il mega download
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "Cap 1: Forma Giur.", 
        "Cap 2: Territorio", 
        "Cap 3: Economico", 
        "Cap 4: Patrimoniale", 
        "Cap 5: Finanziario", 
        "Cap 6: Benchmark", 
        "Cap 7: Ranking", 
        "⭐ Scarica Tutto"
    ])

    # --- SCHEDA CAPITOLO 1 ---
    with tab1:
        st.subheader("1. Analisi Forma Giuridica")
        st.write("Genera l'analisi aggregata per S.p.A. e S.r.l.")
        if st.button("Genera Capitolo 1", type="primary", key="btn_cap1"):
            with st.spinner("Creazione tabelle in corso..."):
                excel_cap1 = elabora_capitolo_1(df_orbis)
                st.download_button(
                    label="📥 Scarica '1. Forma Giuridica'",
                    data=excel_cap1,
                    file_name="1_Forma_Giuridica.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # --- SCHEDA CAPITOLO 2 ---
    with tab2:
        st.subheader("2. Analisi Territoriale")
        st.write("Genera il report aggregato per NUTS2 (Regioni e Macroregioni) con grafici a torta e istogrammi.")
        if st.button("Genera Capitolo 2", type="primary", key="btn_cap2"):
            with st.spinner("Creazione tabelle e grafici territoriali..."):
                excel_cap2 = elabora_capitolo_2(df_orbis)
                st.download_button(
                    label="📥 Scarica '2. Ripartizione Territoriale'",
                    data=excel_cap2,
                    file_name="2_Ripartizione_Territoriale.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # --- SCHEDA CAPITOLO 3 ---
    with tab3:
        st.subheader("3. Equilibrio Economico & Svil. Dimensionale")
        st.write("Genera analisi approfondite su Margini (Profitto, EBITDA, EBIT) e Ricavi.")
        if st.button("Genera Capitolo 3", type="primary", key="btn_cap3"):
            with st.spinner("Calcolo indici economici..."):
                excel_cap3 = elabora_capitolo_3(df_orbis)
                st.download_button(
                    label="📥 Scarica '3. Equilibrio Economico'",
                    data=excel_cap3,
                    file_name="3_Equilibrio_Economico.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # --- SCHEDA CAPITOLO 4 ---
    with tab4:
        st.subheader("4. Equilibrio Patrimoniale")
        st.write("Genera analisi sugli indici di Struttura e Gearing.")
        if st.button("Genera Capitolo 4", type="primary", key="btn_cap4"):
            with st.spinner("Calcolo metriche patrimoniali..."):
                excel_cap4 = elabora_capitolo_4(df_orbis)
                st.download_button(
                    label="📥 Scarica '4. Equilibrio Patrimoniale'",
                    data=excel_cap4,
                    file_name="4_Equilibrio_Patrimoniale.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # --- SCHEDA CAPITOLO 5 ---
    with tab5:
        st.subheader("5. Equilibrio Finanziario")
        st.write("Genera analisi sugli indici di liquidità (Current, Quick) e rotazione.")
        if st.button("Genera Capitolo 5", type="primary", key="btn_cap5"):
            with st.spinner("Calcolo metriche finanziarie..."):
                excel_cap5 = elabora_capitolo_5(df_orbis)
                st.download_button(
                    label="📥 Scarica '5. Equilibrio Finanziario'",
                    data=excel_cap5,
                    file_name="5_Equilibrio_Finanziario.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # --- SCHEDA CAPITOLO 6 ---
    with tab6:
        st.subheader("6. Rating e Benchmark")
        st.write("Genera il cruscotto finale con calcolo terzili, assegnazione rating A-B-C e pivot territoriali.")
        if st.button("Genera Capitolo 6", type="primary", key="btn_cap6"):
            with st.spinner("Calcolo benchmark e assegnazione rating..."):
                excel_cap6 = elabora_capitolo_6(df_orbis)
                st.download_button(
                    label="📥 Scarica '6. Benchmark'",
                    data=excel_cap6,
                    file_name="6_Benchmark.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # --- SCHEDA CAPITOLO 7 ---
    with tab7:
        st.subheader("7. Ranking Aziendale")
        st.write("Elabora le classifiche nazionali e regionali. Scaricherai un archivio ZIP contenente 3 file Excel.")
        if st.button("Genera Capitolo 7 (Pack ZIP)", type="primary", key="btn_cap7"):
            with st.spinner("Creazione ranking e compressione file..."):
                zip_cap7 = elabora_capitolo_7(df_orbis)
                st.download_button(
                    label="📥 Scarica '7. Pacchetto Ranking' (ZIP)",
                    data=zip_cap7,
                    file_name="7_Ranking_Aziendale_Pack.zip",
                    mime="application/zip"
                )

    # --- SCHEDA SCARICA TUTTO ---
    with tab8:
        st.subheader("⭐ Master Export: Tutti i Capitoli")
        st.write("Con un solo clic, Python elaborerà tutti e 7 i capitoli e ti restituirà un unico archivio ZIP contenente l'intero progetto.")
        if st.button("🚀 GENERA INTERO PROGETTO", type="primary", use_container_width=True, key="btn_all"):
            with st.spinner("Elaborazione massiva in corso... Mettiti comodo, potrebbe volerci qualche secondo!"):
                
                # Creiamo il mega ZIP in memoria
                master_zip_buffer = io.BytesIO()
                import zipfile # Assicurati sia importato
                
                with zipfile.ZipFile(master_zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as master_zip:
                    # Eseguiamo e salviamo i primi 6 capitoli
                    master_zip.writestr("1_Forma_Giuridica.xlsx", elabora_capitolo_1(df_orbis).read())
                    master_zip.writestr("2_Ripartizione_Territoriale.xlsx", elabora_capitolo_2(df_orbis).read())
                    master_zip.writestr("3_Equilibrio_Economico.xlsx", elabora_capitolo_3(df_orbis).read())
                    master_zip.writestr("4_Equilibrio_Patrimoniale.xlsx", elabora_capitolo_4(df_orbis).read())
                    master_zip.writestr("5_Equilibrio_Finanziario.xlsx", elabora_capitolo_5(df_orbis).read())
                    master_zip.writestr("6_Benchmark.xlsx", elabora_capitolo_6(df_orbis).read())
                    
                    # Il Capitolo 7 è già uno ZIP. Lo estraiamo e mettiamo i 3 file nel Mega ZIP
                    cap7_zip_buffer = elabora_capitolo_7(df_orbis)
                    with zipfile.ZipFile(cap7_zip_buffer, "r") as cap7_zip:
                        for nome_file in cap7_zip.namelist():
                            master_zip.writestr(nome_file, cap7_zip.read(nome_file))

                master_zip_buffer.seek(0)
                
                st.success("Tutti i capitoli elaborati con successo!")
                st.download_button(
                    label="📥 SCARICA PROGETTO COMPLETO (.zip)",
                    data=master_zip_buffer,
                    file_name="FinHack_Report_Completo.zip",
                    mime="application/zip",
                    use_container_width=True
                )
