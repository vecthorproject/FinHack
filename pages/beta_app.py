import streamlit as st
import pandas as pd
import numpy as np
import io
import xlsxwriter.utility as xl_util

st.set_page_config(page_title="FinHack - Beta Totale", page_icon="📈", layout="wide")

st.title("📈 FinHack: Calcolo Beta Totale")
st.markdown("Carica il file di estrazione di ORBIS e l'output della regressione per generare il modello completo.")
st.divider()

st.info("💡 **ISTRUZIONI:** Carica l'output regressione **BTOT-SW** e un export di **ORBIS** in formato `.xlsx`. Assicurati di aver usato i filtri corretti e il formato **LISTA BETA TOTALE**.")

# --- INIZIALIZZAZIONE SESSION STATE SICURA ---
if "file_generated" not in st.session_state:
    st.session_state.file_generated = False
if "excel_data" not in st.session_state:
    st.session_state.excel_data = None
if "last_orbis" not in st.session_state:
    st.session_state.last_orbis = None
if "last_reg" not in st.session_state:
    st.session_state.last_reg = None

# --- 1. UPLOAD DEI FILE ---
col1, col2 = st.columns(2)
with col1:
    orbis_file = st.file_uploader("Carica l'export di Orbis (Excel)", type=["xlsx"])
with col2:
    reg_file = st.file_uploader("Carica il file Regressioni Beta (Excel o CSV)", type=["xlsx", "csv"])

# Reset dello stato in caso di caricamento di nuovi file
if orbis_file and reg_file:
    if st.session_state.last_orbis != orbis_file.name or st.session_state.last_reg != reg_file.name:
        st.session_state.excel_data = None
        st.session_state.file_generated = False
        st.session_state.last_orbis = orbis_file.name
        st.session_state.last_reg = reg_file.name

def clean_italian_number(x):
    if pd.isna(x): return np.nan
    val_str = str(x).strip().upper()
    if val_str in ['ND', 'N.D.', 'NA', '', 'N.A.']: return np.nan
    if isinstance(x, str):
        x = x.replace('.', '').replace(',', '.')
    try:
        return float(x)
    except:
        return np.nan

def get_col_names(df, keyword):
    return [c for c in df.columns if keyword.lower() in str(c).lower()]

# --- 2. ELABORAZIONE E COSTRUZIONE MODELLO ---
if orbis_file and reg_file and not st.session_state.file_generated:
    with st.spinner("Mappatura formule e allineamento coefficienti nell'Excel... ⏳"):
        try:
            # Lettura dei fogli grezzi mantenendo l'esatta origine dei dati
            df_orbis_raw = pd.read_excel(orbis_file, sheet_name="Risultati", engine='calamine')
            
            if reg_file.name.endswith('.csv'):
                df_reg_raw = pd.read_csv(reg_file, sep=',', encoding='utf-8')
            else:
                try:
                    df_reg_raw = pd.read_excel(reg_file, sheet_name="BTOT-SW", engine='calamine')
                except Exception:
                    df_reg_raw = pd.read_excel(reg_file, engine='calamine')

            df_orbis = df_orbis_raw.copy()
            df_reg = df_reg_raw.copy()

            # Allineamento dei nomi delle colonne per il calcolo del filtro in Python
            col_mapping_reg = {
                'Settore': 'Settore', 'Interc': 'coeff Intercetta', 'VP EBIT': 'coeff VP EBIT',
                'Roe': 'coeff Roe', 'ln Fatturato': 'coeff ln Fatturato', 'CCPG/Attivo': 'coeff CCPG/Attivo',
                'Leverage': 'coeff Leverage', 'Current ratio': 'coeff Current ratio', 'Interest coverage': 'coeff Interest coverage'
            }
            df_reg = df_reg.rename(columns=col_mapping_reg)
            coeff_cols = [c for c in col_mapping_reg.values() if c != 'Settore']
            
            for c in coeff_cols:
                if c in df_reg.columns:
                    df_reg[c] = df_reg[c].fillna(0).astype(float)

            # Identificazione delle colonne di Orbis
            col_nome = get_col_names(df_orbis, "ragione sociale")
            col_bvd = get_col_names(df_orbis, "bvd id")
            col_settore = get_col_names(df_orbis, "bvd sectors")
            
            nome_key = col_nome[0] if col_nome else df_orbis.columns[0]
            bvd_key = col_bvd[0] if col_bvd else df_orbis.columns[1]
            settore_key = col_settore[0] if col_settore else df_orbis.columns[2]

            cols_val_prod = get_col_names(df_orbis, "valore della produzione")
            cols_vp_ebit = get_col_names(df_orbis, "vp ebit")
            cols_roe = get_col_names(df_orbis, "roe")
            cols_ccn = get_col_names(df_orbis, "ccn/tot ass")
            cols_lev = get_col_names(df_orbis, "lev (")
            cols_cr = get_col_names(df_orbis, "cr (")
            cols_ic = get_col_names(df_orbis, "ic (")

            def get_excel_col_letter(df, col_name):
                return xl_util.xl_col_to_name(df.columns.get_loc(col_name))

            let_nome = get_excel_col_letter(df_orbis, nome_key)
            let_bvd = get_excel_col_letter(df_orbis, bvd_key)
            let_settore = get_excel_col_letter(df_orbis, settore_key)

            let_vp_ebit = [get_excel_col_letter(df_orbis, c) for c in cols_vp_ebit]
            let_roe = [get_excel_col_letter(df_orbis, c) for c in cols_roe]
            let_ccn = [get_excel_col_letter(df_orbis, c) for c in cols_ccn]
            let_lev = [get_excel_col_letter(df_orbis, c) for c in cols_lev]
            let_cr = [get_excel_col_letter(df_orbis, c) for c in cols_cr]
            let_ic = [get_excel_col_letter(df_orbis, c) for c in cols_ic]
            let_val_prod = [get_excel_col_letter(df_orbis, c) for c in cols_val_prod]

            # Pulizia e calcolo temporaneo in Python per generare la maschera delle righe valide
            df_calc_py = pd.DataFrame()
            df_calc_py['Settore'] = df_orbis[settore_key]
            for c in cols_val_prod + cols_vp_ebit + cols_roe + cols_ccn + cols_lev + cols_cr + cols_ic:
                df_orbis[c] = df_orbis[c].apply(clean_italian_number)

            def calc_ln_mean(row):
                vals = pd.to_numeric(row, errors='coerce').dropna()
                vals = vals[vals > 0]
                if len(vals) == 0: return np.nan
                return np.log(vals * 1000).mean()

            py_lnfatt = df_orbis[cols_val_prod].apply(calc_ln_mean, axis=1)
            py_vp_ebit = df_orbis[cols_vp_ebit].mean(axis=1)
            py_roe = df_orbis[cols_roe].mean(axis=1)
            
            df_final_py = pd.merge(df_calc_py, df_reg[['Settore'] + coeff_cols], on='Settore', how='left')
            
            py_beta = (
                df_final_py['coeff Intercetta'] + (df_final_py['coeff VP EBIT'] * py_vp_ebit) +
                (df_final_py['coeff Roe'] * py_roe) + (df_final_py['coeff ln Fatturato'] * py_lnfatt) +
                (df_final_py['coeff CCPG/Attivo'] * df_orbis[cols_ccn].mean(axis=1)) + (df_final_py['coeff Leverage'] * df_orbis[cols_lev].mean(axis=1)) +
                (df_final_py['coeff Current ratio'] * df_orbis[cols_cr].mean(axis=1)) + (df_final_py['coeff Interest coverage'] * df_orbis[cols_ic].mean(axis=1))
            )
            
            valid_rows_indices = df_orbis[py_beta.notna() & ~np.isinf(py_beta) & py_vp_ebit.notna() & py_roe.notna()].index.tolist()

            # --- 3. SCRITTURA WORKBOOK EXCEL ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                format_dec = workbook.add_format({'num_format': '0.0000', 'border': 1, 'border_color': '#D9D9D9'})
                format_str = workbook.add_format({'border': 1, 'border_color': '#D9D9D9'})
                
                header_orbis = workbook.add_format({'bold': True, 'bg_color': '#B4C6E7', 'font_color': 'black', 'border': 1})
                header_coeff = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'font_color': 'black', 'border': 1})
                header_beta  = workbook.add_format({'bold': True, 'bg_color': '#FFC000', 'font_color': 'black', 'border': 2})
                
                cell_coeff = workbook.add_format({'num_format': '0.0000', 'bg_color': '#F8FBF8', 'border': 1, 'border_color': '#D9D9D9'})
                cell_beta  = workbook.add_format({'num_format': '0.0000', 'bg_color': '#FFF2CC', 'border': 2, 'bold': True})

                df_orbis_raw.to_excel(writer, index=False, sheet_name='Risultati')
                df_reg_raw.to_excel(writer, index=False, sheet_name='BTOT-SW')

                visible_columns = [
                    'Nome', 'Settore', 'BVB ID', 'VP EBIT', 'ROE', 'LN FATT', 
                    'CCN/TOT ASS', 'LEV', 'CR', 'IC', 
                    'coeff Intercetta', 'coeff VP EBIT', 'coeff Roe', 'coeff ln Fatturato', 
                    'coeff CCPG/Attivo', 'coeff Leverage', 'coeff Current ratio', 'coeff Interest coverage', 
                    'Beta Totale'
                ]

                def write_calc_sheet(sheet_name, rows_to_include):
                    worksheet = workbook.add_worksheet(sheet_name)
                    
                    for col_num, col_name in enumerate(visible_columns):
                        if 'coeff' in col_name:
                            worksheet.write(0, col_num, col_name, header_coeff)
                            worksheet.set_column(col_num, col_num, 15)
                        elif col_name == 'Beta Totale':
                            worksheet.write(0, col_num, col_name, header_beta)
                            worksheet.set_column(col_num, col_num, 18)
                        else:
                            worksheet.write(0, col_num, col_name, header_orbis)
                            worksheet.set_column(col_num, col_num, 25 if col_num < 3 else 16)

                    for new_row_idx, original_idx in enumerate(rows_to_include):
                        r_calc = new_row_idx + 2
                        r_orig = original_idx + 2
                        
                        worksheet.write_formula(new_row_idx + 1, 0, f"=Risultati!{let_nome}{r_orig}", format_str)
                        worksheet.write_formula(new_row_idx + 1, 1, f"=Risultati!{let_settore}{r_orig}", format_str)
                        worksheet.write_formula(new_row_idx + 1, 2, f"=Risultati!{let_bvd}{r_orig}", format_str)
                        
                        worksheet.write_formula(new_row_idx + 1, 3, f"=AVERAGE(Risultati!{let_vp_ebit[0]}{r_orig}:{let_vp_ebit[-1]}{r_orig})", format_dec)
                        worksheet.write_formula(new_row_idx + 1, 4, f"=AVERAGE(Risultati!{let_roe[0]}{r_orig}:{let_roe[-1]}{r_orig})", format_dec)
                        
                        sum_ln_parts = " + ".join([f"IF(Risultati!{let}{r_orig}>0, LN(Risultati!{let}{r_orig}*1000), 0)" for let in let_val_prod])
                        count_ln_parts = " + ".join([f"IF(Risultati!{let}{r_orig}>0, 1, 0)" for let in let_val_prod])
                        formula_ln = f"=IF(({count_ln_parts})>0, ({sum_ln_parts}) / ({count_ln_parts}), NA())"
                        worksheet.write_formula(new_row_idx + 1, 5, formula_ln, format_dec)
                        
                        worksheet.write_formula(new_row_idx + 1, 6, f"=AVERAGE(Risultati!{let_ccn[0]}{r_orig}:{let_ccn[-1]}{r_orig})", format_dec)
                        worksheet.write_formula(new_row_idx + 1, 7, f"=AVERAGE(Risultati!{let_lev[0]}{r_orig}:{let_lev[-1]}{r_orig})", format_dec)
                        worksheet.write_formula(new_row_idx + 1, 8, f"=AVERAGE(Risultati!{let_cr[0]}{r_orig}:{let_cr[-1]}{r_orig})", format_dec)
                        worksheet.write_formula(new_row_idx + 1, 9, f"=AVERAGE(Risultati!{let_ic[0]}{r_orig}:{let_ic[-1]}{r_orig})", format_dec)

                        # --- CORREZIONE INDICI VLOOKUP (Puntano ai Coefficienti Reali, colonne pari) ---
                        vlookup_base = f"=VLOOKUP(B{r_calc}, 'BTOT-SW'!$B$2:$V$200"
                        worksheet.write_formula(new_row_idx + 1, 10, f"{vlookup_base}, 6, FALSE)", cell_coeff)  # Intercetta (Col G)
                        worksheet.write_formula(new_row_idx + 1, 11, f"{vlookup_base}, 20, FALSE)", cell_coeff) # VP EBIT (Col U)
                        worksheet.write_formula(new_row_idx + 1, 12, f"{vlookup_base}, 8, FALSE)", cell_coeff)  # Roe (Col I)
                        worksheet.write_formula(new_row_idx + 1, 13, f"{vlookup_base}, 10, FALSE)", cell_coeff) # ln Fatturato (Col K)
                        worksheet.write_formula(new_row_idx + 1, 14, f"{vlookup_base}, 12, FALSE)", cell_coeff) # CCPG/Attivo (Col M)
                        worksheet.write_formula(new_row_idx + 1, 15, f"{vlookup_base}, 14, FALSE)", cell_coeff) # Leverage (Col O)
                        worksheet.write_formula(new_row_idx + 1, 16, f"{vlookup_base}, 16, FALSE)", cell_coeff) # Current ratio (Col Q)
                        worksheet.write_formula(new_row_idx + 1, 17, f"{vlookup_base}, 18, FALSE)", cell_coeff) # Interest coverage (Col S)
                        
                        formula_beta = f"=K{r_calc} + (L{r_calc}*D{r_calc}) + (M{r_calc}*E{r_calc}) + (N{r_calc}*F{r_calc}) + (O{r_calc}*G{r_calc}) + (P{r_calc}*H{r_calc}) + (Q{r_calc}*I{r_calc}) + (R{r_calc}*J{r_calc})"
                        worksheet.write_formula(new_row_idx + 1, 18, formula_beta, cell_beta)

                write_calc_sheet('Calcolo Beta Tot', list(range(len(df_orbis))))
                write_calc_sheet('Calcolo Beta Tot FIX', valid_rows_indices)

            st.session_state.excel_data = output.getvalue()
            st.session_state.file_generated = True

        except Exception as e:
            st.error(f"Errore durante l'elaborazione. Verifica la struttura dei fogli. Errore: {e}")

# --- 4. INTERFACCIA DI DOWNLOAD PULITA ---
if st.session_state.file_generated and st.session_state.excel_data:
    st.success("✅ File di Simulazione Calcolo Beta Totale generato con successo!")
    st.download_button(
        label="📥 Scarica Calc.BetaTot",
        data=st.session_state.excel_data,
        file_name="Calcolo_BetaTot.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
