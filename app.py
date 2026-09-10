import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stats Fantacalcio", layout="centered")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@st.cache_data
def carica_dati():
    file_path = os.path.join(BASE_DIR, 'FANTA APP.xlsx')
    
    if not os.path.exists(file_path):
        for f in os.listdir(BASE_DIR):
            if f.lower() == 'fanta app.xlsx':
                file_path = os.path.join(BASE_DIR, f)
                break

    if not os.path.exists(file_path):
        return pd.DataFrame()

    try:
        xls = pd.ExcelFile(file_path)
        
        # 1. LETTURA DEL FOGLIO PRINCIPALE
        df = pd.read_excel(file_path, sheet_name=xls.sheet_names[0])
        df.columns = [str(c).strip() for c in df.columns]
        
        col_nome = next((c for c in df.columns if str(c).lower() in ['nome', 'nome giocatore', 'giocatore']), None)
        
        if not col_nome:
            df = pd.read_excel(file_path, sheet_name=xls.sheet_names[0], header=1)
            df.columns = [str(c).strip() for c in df.columns]
            col_nome = next((c for c in df.columns if str(c).lower() in ['nome', 'nome giocatore', 'giocatore']), None)
            
        if not col_nome:
            for col in df.columns:
                if df[col].astype(str).str.contains('[a-zA-Z]').any():
                    col_nome = col
                    break

        if col_nome:
            df['Nome_Clean'] = df[col_nome].astype(str).str.strip()
            df = df.dropna(subset=['Nome_Clean'])
            df = df[~df['Nome_Clean'].str.upper().isin(['NOME', 'NOME GIOCATORE', 'TOP', 'MEDI', 'LOW', 'NAN', ''])]
            df = df[~df['Nome_Clean'].str.isnumeric()]
            df['Nome_Match'] = df['Nome_Clean'].str.lower()
        else:
            return pd.DataFrame()

        # 2. RICERCA ESTRAZIONE PREZZI DA *TUTTI* GLI ALTRI FOGLI
        list_prezzi = []
        for sheet in xls.sheet_names:
            df_s = pd.read_excel(file_path, sheet_name=sheet)
            
            # Cerca intestazioni sepolte nelle righe
            header_mask = df_s.astype(str).apply(lambda col: col.str.contains('Nome Giocatore|Nome', case=False, na=False)).any(axis=1)
            idx_list = header_mask[header_mask].index.tolist()
            
            if not idx_list:
                c_n = next((c for c in df_s.columns if str(c).lower() in ['nome', 'nome giocatore', 'giocatore']), None)
                if not c_n:
                    df_s = pd.read_excel(file_path, sheet_name=sheet, header=1)
                idx_list = [0]

            for idx in idx_list:
                if idx == 0 and not header_mask.get(0, False):
                    sub = df_s.copy()
                else:
                    sub = df_s.iloc[idx + 1:].copy()
                    sub.columns = df_s.iloc[idx].astype(str).str.strip()
                
                c_nome = next((c for c in sub.columns if str(c).lower() in ['nome', 'nome giocatore', 'giocatore']), None)
                if c_nome:
                    p_max_col = next((c for c in sub.columns if 'max' in str(c).lower()), None)
                    p_min_col = next((c for c in sub.columns if 'min' in str(c).lower()), None)
                    p_med_col = next((c for c in sub.columns if 'medio' in str(c).lower()), None)
                    
                    if p_max_col or p_min_col or p_med_col:
                        temp = sub[[c_nome]].copy()
                        temp['Nome_Match'] = temp[c_nome].astype(str).str.strip().str.lower()
                        if p_max_col: temp['P_MAX_estratto'] = pd.to_numeric(sub[p_max_col], errors='coerce')
                        if p_min_col: temp['P_MIN_estratto'] = pd.to_numeric(sub[p_min_col], errors='coerce')
                        if p_med_col: temp['P_MED_estratto'] = pd.to_numeric(sub[p_med_col], errors='coerce')
                        
                        temp = temp.dropna(subset=['Nome_Match'])
                        list_prezzi.append(temp)
                        
        # 3. UNIONE TOTALE DATI
        if list_prezzi:
            df_prezzi = pd.concat(list_prezzi, ignore_index=True)
            df_prezzi = df_prezzi.groupby('Nome_Match').first().reset_index()
            df = pd.merge(df, df_prezzi, on='Nome_Match', how='left')

        return df

    except Exception as e:
        return pd.DataFrame()

df = carica_dati()

# --- INTERFACCIA UTENTE ---
st.title('⚽ Dashboard Fantacalcio & Asta')

if not df.empty and 'Nome_Clean' in df.columns:
    lista_giocatori = sorted(df['Nome_Clean'].unique().tolist())
    lista_giocatori = [g for g in lista_giocatori if str(g).lower() != 'nan' and g != '']
    lista_giocatori.insert(0, '')

    giocatore_scelto = st.selectbox(
        'Scrivi o seleziona il nome del giocatore:', lista_giocatori
    )

    if giocatore_scelto:
        dati = df[df['Nome_Clean'] == giocatore_scelto].iloc[0]

        def ottieni_valore(*nomi_colonne, default=0):
            for col in nomi_colonne:
                col_cercata = str(col).replace(" ", "").replace(".", "").lower()
                
                match_col = next((c for c in dati.index if str(c).replace(" ", "").replace(".", "").lower() == col_cercata), None)
                if not match_col:
                    match_col = next((c for c in dati.index if col_cercata in str(c).replace(" ", "").replace(".", "").lower()), None)

                if match_col and pd.notna(dati[match_col]):
                    val = dati[match_col]
                    val_str = str(val).strip().upper()
                    if val_str not in ['', 'NAN', 'N.D.', 'NULL', 'NONE']:
                        return val
            return default

        nome_display = ottieni_valore('Nome', 'Nome Giocatore', 'Giocatore', default=giocatore_scelto)
        ruolo = ottieni_valore('R', 'Ruolo', default='N.D.')
        squadra = ottieni_valore('Squadra', default='N.D.')

        st.subheader(f"👤 {nome_display}")
        st.write(f"**Ruolo:** {ruolo} | **Squadra:** {squadra}")
        st.divider()

        fvm = ottieni_valore('FVM', 'Consigliato', 'FVM 26/27', default='N.D.')
        qt_i = ottieni_valore('Qt.I', 'Iniziale', 'Qt Iniziale', 'Quotazione', default='N.D.')
        qt_a = ottieni_valore('Qt.A', 'Attuale', 'Qt Attuale', default='N.D.')
        
        # Peschiamo prioritariamente i dati estratti dai fogli aggiuntivi
        p_max = ottieni_valore('P_MAX_estratto', 'Prezzo Max', 'Max', default=0)
        p_min = ottieni_valore('P_MIN_estratto', 'Prezzo Min', 'Min', default=0)
        p_med = ottieni_valore('P_MED_estratto', 'Prezzo Medio', 'Medio', default=0)

        def pulisci_numero(val):
            try:
                return float(str(val).replace(',', '.'))
            except:
                return 0.0

        val_qta = pulisci_numero(qt_a)
        val_qti = pulisci_numero(qt_i)
        
        fallback_price = val_qta if val_qta > 0 else val_qti

        val_pmax = pulisci_numero(p_max) if pulisci_numero(p_max) > 0 else fallback_price
        val_pmin = pulisci_numero(p_min) if pulisci_numero(p_min) > 0 else fallback_price
        val_pmed = pulisci_numero(p_med) if pulisci_numero(p_med) > 0 else fallback_price

        p_max_str = f"{int(val_pmax)} cr." if val_pmax > 0 else "0 cr."
        p_min_str = f"{int(val_pmin)} cr." if val_pmin > 0 else "N.D."
        p_med_str = f"{int(round(val_pmed))} cr." if val_pmed > 0 else "0 cr."

        col_a1, col_a2, col_a3, col_a4 = st.columns(4)
        col_a1.metric('Prezzo Max', p_max_str)
        col_a2.metric('Prezzo Min', p_min_str)
        col_a3.metric('Prezzo Medio', p_med_str)
        col_a4.metric('FVM Consigliato', str(fvm))

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader('📊 Statistiche / Storico')
            st.write(f"**Presenze:** {ottieni_valore('Presenze', 'Pv', 'P', default=0)}")
            st.write(f"**Media Voto (MV):** {ottieni_valore('MV', 'Media Voto', default=0.0)}")
            st.write(f"**Fantamedia (FM):** {ottieni_valore('FM', 'Fantamedia', default=0.0)}")
            st.write(f"**Gol Fatti:** {ottieni_valore('Gol', 'Gf', default=0)} | **Assist:** {ottieni_valore('Assist', 'Ass', default=0)}")
            st.write(f"**Ammonizioni:** {ottieni_valore('Ammonizioni', 'Amm', 'Au', default=0)}")

        with col2:
            st.subheader('💰 Quotazioni 2026/2027')
            st.write(f"**Quotazione Iniziale (Qt.I):** {qt_i}")
            st.write(f"**Quotazione Attuale (Qt.A):** {qt_a}")
else:
    st.error('File "FANTA APP.xlsx" non trovato o impossibile leggere i nomi. Assicurati di aver fatto il Reboot dell\'app.')