import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stats Fantacalcio", layout="centered")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@st.cache_data
def carica_dati():
    file_path = os.path.join(BASE_DIR, 'FANTA APP.xlsx')
    
    if not os.path.exists(file_path):
        return pd.DataFrame()

    try:
        # Legge il primo foglio del tuo file unico
        xls = pd.ExcelFile(file_path)
        df = pd.read_excel(file_path, sheet_name=xls.sheet_names[0])
        
        # Pulisce i nomi delle colonne da eventuali spazi extra
        df.columns = [str(c).strip() for c in df.columns]
        
        # Trova automaticamente la colonna del nome (Nome, Nome Giocatore, ecc.)
        col_nome = next((c for c in df.columns if str(c).lower() in ['nome', 'nome giocatore', 'giocatore']), None)
        
        if col_nome:
            df['Nome_Clean'] = df[col_nome].astype(str).str.strip()
            df = df.dropna(subset=['Nome_Clean'])
            # Filtra eventuali righe di intestazione interne residue
            df = df[~df['Nome_Clean'].str.upper().isin(['NOME', 'NOME GIOCATORE', 'TOP', 'MEDI', 'LOW', 'NAN'])]
        
        return df

    except Exception as e:
        st.error(f"Errore nella lettura del file: {e}")
        return pd.DataFrame()

df = carica_dati()

# --- INTERFACCIA UTENTE ---
st.title('⚽ Dashboard Fantacalcio & Asta')

if not df.empty and 'Nome_Clean' in df.columns:
    lista_giocatori = sorted(df['Nome_Clean'].unique().tolist())
    lista_giocatori.insert(0, '')

    giocatore_scelto = st.selectbox(
        'Scrivi o seleziona il nome del giocatore:', lista_giocatori
    )

    if giocatore_scelto:
        dati = df[df['Nome_Clean'] == giocatore_scelto].iloc[0]

        # Funzione corazzata: cerca il valore testando tutte le possibili varianti del nome colonna
        def ottieni_valore(*nomi_colonne):
            for col in nomi_colonne:
                if col in dati and pd.notna(dati[col]):
                    val = dati[col]
                    if str(val).strip() not in ['', 'nan', 'N.D.']:
                        return val
            return 0

        st.subheader(f"👤 {giocatore_scelto}")
        ruolo = ottieni_valore('R', 'Ruolo', 'Ruolo 26/27')
        squadra = ottieni_valore('Squadra', 'Squadra 26/27')
        
        st.write(f"**Ruolo:** {ruolo if ruolo != 0 else 'N.D.'} | **Squadra:** {squadra if squadra != 0 else 'N.D.'}")
        st.divider()

        # Prezzi e Valutazioni
        p_max = ottieni_valore('Prezzo Max', 'Max')
        p_min = ottieni_valore('Prezzo Min', 'Min')
        p_med = ottieni_valore('Prezzo Medio', 'Medio')
        fvm = ottieni_valore('FVM', 'FVM 26/27', 'FVM Consigliato')

        col_a1, col_a2, col_a3, col_a4 = st.columns(4)
        col_a1.metric('Prezzo Max', f"{int(p_max)} cr." if p_max != 0 else "0 cr.")
        col_a2.metric('Prezzo Min', f"{int(p_min)} cr." if p_min != 0 else "N.D.")
        col_a3.metric('Prezzo Medio', f"{int(float(p_med))} cr." if p_med != 0 else "0 cr.")
        col_a4.metric('FVM Consigliato', f"{fvm}" if fvm != 0 else "N.D.")

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader('📊 Statistiche / Storico')
            st.write(f"**Presenze:** {ottieni_valore('Presenze', 'Pv', 'P', 'Presenze 25/26')}")
            st.write(f"**Media Voto (MV):** {ottieni_valore('MV', 'Mv', 'MV 25/26')}")
            st.write(f"**Fantamedia (FM):** {ottieni_valore('FM', 'Fm', 'FM 25/26')}")
            st.write(f"**Gol Fatti:** {ottieni_valore('Gol', 'Gf', 'Gol 25/26')} | **Assist:** {ottieni_valore('Assist', 'Ass', 'Assist 25/26')}")
            st.write(f"**Ammonizioni:** {ottieni_valore('Ammonizioni', 'Amm', 'Ammoniz.', 'Ammoniz. 25/26', 'Au')}")

        with col2:
            st.subheader('💰 Quotazioni 2026/2027')
            qt_i = ottieni_valore('Qt.I', 'Quotazione Iniziale')
            qt_a = ottieni_valore('Qt.A', 'Quotazione Attuale')
            st.write(f"**Quotazione Iniziale (Qt.I):** {qt_i if qt_i != 0 else 'N.D.'}")
            st.write(f"**Quotazione Attuale (Qt.A):** {qt_a if qt_a != 0 else 'N.D.'}")
else:
    st.error('File "FANTA APP.xlsx" non trovato o vuoto. Assicurati che sia presente nel repository con quel nome esatto.')