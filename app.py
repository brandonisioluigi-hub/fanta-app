import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stats Fantacalcio", layout="centered")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@st.cache_data
def carica_dati():
    # 1. Riconoscimento dinamico e sicuro dei 4 file
    file_2526, file_2627, file_tot, file_nuovi = None, None, None, None
    
    for f in os.listdir(BASE_DIR):
        if f.endswith('.xlsx') and not f.startswith('~$'):
            f_lower = f.lower()
            if 'statistiche' in f_lower and '2025_26' in f_lower: file_2526 = f
            elif 'quotazioni' in f_lower: file_2627 = f
            elif 'totale' in f_lower: file_tot = f
            elif 'nuovi_acquisti' in f_lower: file_nuovi = f

    # 2. Caricamento file
    df_tot = pd.read_excel(os.path.join(BASE_DIR, file_tot), sheet_name='Statistiche_Incrociate') if file_tot else pd.DataFrame()
    df_2627 = pd.read_excel(os.path.join(BASE_DIR, file_2627), sheet_name='Tutti', header=1) if file_2627 else pd.DataFrame()
    df_2526 = pd.read_excel(os.path.join(BASE_DIR, file_2526), sheet_name='Tutti', header=1) if file_2526 else pd.DataFrame()
    df_nuovi = pd.read_excel(os.path.join(BASE_DIR, file_nuovi)) if file_nuovi else pd.DataFrame()

    # Logica per recuperare i Prezzi Minimi dalle altre tabelle in Totale.xlsx
    df_prices = pd.DataFrame(columns=['Nome_Clean', 'Prezzo Min'])
    if file_tot:
        xls_tot = pd.ExcelFile(os.path.join(BASE_DIR, file_tot))
        list_subs = []
        for sheet in ['Portieri_Completo', 'Difensori_Completo', 'Centrocampo_Completo', 'Attacco_Completo']:
            if sheet in xls_tot.sheet_names:
                df_s = pd.read_excel(xls_tot, sheet_name=sheet)
                idx_headers = df_s[df_s.iloc[:, 0] == 'Nome Giocatore'].index
                for idx in idx_headers:
                    sub = df_s.iloc[idx+1 :].dropna(subset=[df_s.columns[0]]).copy()
                    sub = sub[~sub.iloc[:, 0].astype(str).str.contains('TOP|MEDI|LOW|🌟', case=False, na=False)]
                    if not sub.empty:
                        sub = sub.iloc[:, :5]
                        sub.columns = ['Nome Giocatore', 'Frequenza', 'Prezzo Max', 'Prezzo Medio', 'Prezzo Min']
                        list_subs.append(sub)
        if list_subs:
            dp = pd.concat(list_subs, ignore_index=True)
            dp['Nome_Clean'] = dp['Nome Giocatore'].astype(str).str.strip().str.lower()
            df_prices = dp.groupby('Nome_Clean')['Prezzo Min'].min().reset_index()

    # 3. Uniformiamo nomi e ID per collegare tutto
    if not df_tot.empty: df_tot['Nome_Clean'] = df_tot['Nome Giocatore'].astype(str).str.strip().str.lower()
    if not df_2627.empty: df_2627['Nome_Clean'] = df_2627['Nome'].astype(str).str.strip().str.lower()

    if 'Id' in df_2627.columns: df_2627['Id'] = pd.to_numeric(df_2627['Id'], errors='coerce')
    if 'Id' in df_2526.columns: df_2526['Id'] = pd.to_numeric(df_2526['Id'], errors='coerce')
    
    if not df_nuovi.empty:
        id_col = [c for c in df_nuovi.columns if str(c).strip().lower() == 'id']
        if id_col: df_nuovi['Id'] = pd.to_numeric(df_nuovi[id_col[0]], errors='coerce')

    # 4. Unione di tutti i file
    merged = pd.merge(df_tot, df_2627[['Nome_Clean', 'Id', 'Qt.A', 'Qt.I']], on='Nome_Clean', how='left')
    merged = pd.merge(merged, df_prices, on='Nome_Clean', how='left')
    
    if not df_2526.empty and 'Id' in df_2526.columns:
        merged = pd.merge(merged, df_2526, on='Id', how='left', suffixes=('', '_2526'))
        
    if not df_nuovi.empty and 'Id' in df_nuovi.columns:
        merged = pd.merge(merged, df_nuovi, on='Id', how='left', suffixes=('', '_nuovi'))

    return merged

df = carica_dati()

# --- INTERFACCIA UTENTE ---
st.title('⚽ Dashboard Fantacalcio & Asta')

if not df.empty and 'Nome Giocatore' in df.columns:
    lista_giocatori = sorted(df['Nome Giocatore'].dropna().unique().tolist())
    lista_giocatori.insert(0, '')

    giocatore_scelto = st.selectbox('Scrivi o seleziona il nome del giocatore:', lista_giocatori)

    if giocatore_scelto:
        dati = df[df['Nome Giocatore'] == giocatore_scelto].iloc[0]

        # LA MAGIA È QUI: Questa funzione cerca il dato in tutte le varianti di colonna. 
        # Se nel file 25/26 è zero o vuoto, pesca in automatico dal file Nuovi Acquisti.
        def ottieni_stat(*colonne_possibili):
            for col in colonne_possibili:
                if col in dati and pd.notna(dati[col]):
                    val = dati[col]
                    if str(val).strip() not in ['', '0', '0.0', 'N.D.', 'nan']:
                        return val
            return 0

        # Mappiamo le statistiche pescando dal file storico O dal file nuovi acquisti
        presenze = ottieni_stat('Pv', 'Presenze 25/26', 'Presenze', 'Presenze_nuovi')
        mv = ottieni_stat('Mv', 'MV 25/26', 'MV', 'MV_nuovi')
        fm = ottieni_stat('Fm', 'FM 25/26', 'FM', 'FM_nuovi')
        gol = ottieni_stat('Gf', 'Gol 25/26', 'Gol', 'Gol_nuovi')
        assist = ottieni_stat('Ass', 'Assist 25/26', 'Assist', 'Assist_nuovi')
        ammonizioni = ottieni_stat('Amm', 'Ammoniz. 25/26', 'Ammonizioni', 'Ammoniz.', 'Ammoniz._nuovi')

        # Interfaccia pulita
        st.subheader(f"👤 {dati.get('Nome Giocatore', giocatore_scelto)}")
        ruolo = dati.get('Ruolo', dati.get('R', 'N.D.'))
        squadra = dati.get('Squadra 26/27', dati.get('Squadra', 'N.D.'))
        st.write(f'**Ruolo:** {ruolo} | **Squadra 26/27:** {squadra}')
        st.divider()

        col_a1, col_a2, col_a3, col_a4 = st.columns(4)
        col_a1.metric('Prezzo Max', f"{dati.get('Prezzo Max', 0)} cr.")
        
        p_min = dati.get('Prezzo Min')
        col_a2.metric('Prezzo Min', f"{int(p_min)} cr." if pd.notna(p_min) else "N.D.")
        
        p_medio = dati.get('Prezzo Medio')
        col_a3.metric('Prezzo Medio', f"{int(round(float(p_medio)))} cr." if pd.notna(p_medio) else "0 cr.")
        
        col_a4.metric('FVM Consigliato', f"{dati.get('FVM', dati.get('FVM 26/27', 'N.D.'))}")

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader('📊 Statistiche / Storico')
            st.write(f"**Presenze:** {presenze}")
            st.write(f"**Media Voto (MV):** {mv}")
            st.write(f"**Fantamedia (FM):** {fm}")
            st.write(f"**Gol Fatti:** {gol} | **Assist:** {assist}")
            st.write(f"**Ammonizioni:** {ammonizioni}")

        with col2:
            st.subheader('💰 Quotazioni 2026/2027')
            st.write(f"**Quotazione Iniziale (Qt.I):** {dati.get('Qt.I', 'N.D.')}")
            st.write(f"**Quotazione Attuale (Qt.A):** {dati.get('Qt.A', 'N.D.')}")
else:
    st.error('Errore: impossibile caricare la base dati. Verifica la presenza dei file.')