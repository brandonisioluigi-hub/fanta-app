import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stats Fantacalcio", layout="centered")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@st.cache_data
def carica_dati():
    file_path = os.path.join(BASE_DIR, 'FANTA APP.xlsx')
    
    # Controllo di sicurezza per case-sensitivity su Streamlit Cloud (Linux)
    if not os.path.exists(file_path):
        for f in os.listdir(BASE_DIR):
            if f.lower() == 'fanta app.xlsx':
                file_path = os.path.join(BASE_DIR, f)
                break

    if not os.path.exists(file_path):
        return pd.DataFrame()

    try:
        xls = pd.ExcelFile(file_path)
        df = pd.read_excel(file_path, sheet_name=xls.sheet_names[0])
        
        df.columns = [str(c).strip() for c in df.columns]
        
        # Cerca la colonna del nome in modo dinamico
        col_nome = next((c for c in df.columns if str(c).lower() in ['nome', 'nome giocatore', 'giocatore']), None)
        
        if col_nome:
            df['Nome_Clean'] = df[col_nome].astype(str).str.strip()
            df = df.dropna(subset=['Nome_Clean'])
            # Filtra eventuali intestazioni ripetute finite nel mezzo
            df = df[~df['Nome_Clean'].str.upper().isin(['NOME', 'NOME GIOCATORE', 'TOP', 'MEDI', 'LOW', 'NAN', ''])]
        else:
            # Se proprio non trova la colonna, usa la prima colonna del foglio
            df['Nome_Clean'] = df.iloc[:, 0].astype(str).str.strip()
        
        return df

    except Exception as e:
        return pd.DataFrame()

df = carica_dati()

# --- INTERFACCIA UTENTE ---
st.title('⚽ Dashboard Fantacalcio & Asta')

if not df.empty and 'Nome_Clean' in df.columns:
    lista_giocatori = sorted(df['Nome_Clean'].unique().tolist())
    lista_giocatori = [g for g in lista_giocatori if g.lower() != 'nan' and g != '']
    lista_giocatori.insert(0, '')

    giocatore_scelto = st.selectbox(
        'Scrivi o seleziona il nome del giocatore:', lista_giocatori
    )

    if giocatore_scelto:
        dati = df[df['Nome_Clean'] == giocatore_scelto].iloc[0]

        # Funzione corazzata: testa le varianti della colonna ignorando maiuscole/minuscole
        def ottieni_valore(*nomi_colonne, default=0):
            for col in nomi_colonne:
                match_col = next((c for c in dati.index if str(c).lower() == str(col).lower()), None)
                if match_col and pd.notna(dati[match_col]):
                    val = dati[match_col]
                    val_str = str(val).strip().upper()
                    if val_str not in ['', 'NAN', 'N.D.', 'NULL', 'NONE']:
                        return val
            return default

        # Dati base
        nome_display = ottieni_valore('Nome', 'Nome Giocatore', 'Giocatore', default=giocatore_scelto)
        ruolo = ottieni_valore('R', 'Ruolo', 'Ruolo 26/27', default='N.D.')
        squadra = ottieni_valore('Squadra', 'Squadra 26/27', default='N.D.')

        st.subheader(f"👤 {nome_display}")
        st.write(f"**Ruolo:** {ruolo} | **Squadra:** {squadra}")
        st.divider()

        # Prezzi
        p_max = ottieni_valore('Prezzo Max', 'Max', default=0)
        p_min = ottieni_valore('Prezzo Min', 'Min', default=0)
        p_med = ottieni_valore('Prezzo Medio', 'Medio', default=0)
        fvm = ottieni_valore('FVM', 'FVM 26/27', 'FVM Consigliato', default='N.D.')

        # Formattazione sicura per evitare errori se l'Excel contiene testo nelle celle dei prezzi
        try:
            p_max_str = f"{int(float(p_max))} cr." if float(p_max) > 0 else "0 cr."
        except:
            p_max_str = "0 cr."
            
        try:
            p_min_str = f"{int(float(p_min))} cr." if float(p_min) > 0 else "N.D."
        except:
            p_min_str = "N.D."
            
        try:
            p_med_str = f"{int(round(float(p_med)))} cr." if float(p_med) > 0 else "0 cr."
        except:
            p_med_str = "0 cr."

        col_a1, col_a2, col_a3, col_a4 = st.columns(4)
        col_a1.metric('Prezzo Max', p_max_str)
        col_a2.metric('Prezzo Min', p_min_str)
        col_a3.metric('Prezzo Medio', p_med_str)
        col_a4.metric('FVM Consigliato', str(fvm))

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader('📊 Statistiche / Storico')
            presenze = ottieni_valore('Presenze', 'Pv', 'P', 'Presenze 25/26', default=0)
            mv = ottieni_valore('MV', 'Mv', 'MV 25/26', default=0.0)
            fm = ottieni_valore('FM', 'Fm', 'FM 25/26', default=0.0)
            gol = ottieni_valore('Gol', 'Gf', 'Gol 25/26', default=0)
            assist = ottieni_valore('Assist', 'Ass', 'Assist 25/26', default=0)
            amm = ottieni_valore('Ammonizioni', 'Amm', 'Ammoniz.', 'Ammoniz. 25/26', 'Au', default=0)
            
            st.write(f"**Presenze:** {presenze}")
            st.write(f"**Media Voto (MV):** {mv}")
            st.write(f"**Fantamedia (FM):** {fm}")
            st.write(f"**Gol Fatti:** {gol} | **Assist:** {assist}")
            st.write(f"**Ammonizioni:** {amm}")

        with col2:
            st.subheader('💰 Quotazioni 2026/2027')
            qt_i = ottieni_valore('Qt.I', 'Quotazione Iniziale', 'Qt. Iniziale', default='N.D.')
            qt_a = ottieni_valore('Qt.A', 'Quotazione Attuale', 'Qt. Attuale', default='N.D.')
            st.write(f"**Quotazione Iniziale (Qt.I):** {qt_i}")
            st.write(f"**Quotazione Attuale (Qt.A):** {qt_a}")
else:
    st.error('File "FANTA APP.xlsx" non trovato o vuoto. Assicurati di aver fatto il Reboot dell\'app su Streamlit.')