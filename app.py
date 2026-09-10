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
        
        # 1. Lettura base
        df = pd.read_excel(file_path, sheet_name=xls.sheet_names[0])
        df.columns = [str(c).strip() for c in df.columns]
        col_nome = next((c for c in df.columns if str(c).lower() in ['nome', 'nome giocatore', 'giocatore']), None)
        
        # 2. Se non trova il nome, legge dalla seconda riga (header=1)
        if not col_nome:
            df = pd.read_excel(file_path, sheet_name=xls.sheet_names[0], header=1)
            df.columns = [str(c).strip() for c in df.columns]
            col_nome = next((c for c in df.columns if str(c).lower() in ['nome', 'nome giocatore', 'giocatore']), None)
            
        # 3. Se ancora non lo trova, intercetta la prima colonna testuale
        if not col_nome:
            for col in df.columns:
                if df[col].astype(str).str.contains('[a-zA-Z]').any():
                    col_nome = col
                    break

        if col_nome:
            df['Nome_Clean'] = df[col_nome].astype(str).str.strip()
            df = df.dropna(subset=['Nome_Clean'])
            # Pulizia intestazioni sporche
            df = df[~df['Nome_Clean'].str.upper().isin(['NOME', 'NOME GIOCATORE', 'TOP', 'MEDI', 'LOW', 'NAN', ''])]
            df = df[~df['Nome_Clean'].str.isnumeric()]
        else:
            return pd.DataFrame()
            
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

        # FUNZIONE SUPER-CORAZZATA: Elimina spazi e punti per trovare la colonna a prescindere da come è scritta in Excel
        def ottieni_valore(*nomi_colonne, default=0):
            for col in nomi_colonne:
                col_cercata = str(col).replace(" ", "").replace(".", "").lower()
                
                # Cerca corrispondenza esatta "ripulita"
                match_col = next((c for c in dati.index if str(c).replace(" ", "").replace(".", "").lower() == col_cercata), None)
                
                # Se non la trova, cerca se è almeno una sottostringa (es. se cerchi "fvm" trova "FVM Consigliato")
                if not match_col:
                    match_col = next((c for c in dati.index if col_cercata in str(c).replace(" ", "").replace(".", "").lower()), None)

                if match_col and pd.notna(dati[match_col]):
                    val = dati[match_col]
                    val_str = str(val).strip().upper()
                    # Restituisce il valore ignorando solo i reali campi vuoti (ma tenendo gli zeri)
                    if val_str not in ['', 'NAN', 'N.D.', 'NULL', 'NONE']:
                        return val
            return default

        nome_display = ottieni_valore('Nome', 'Nome Giocatore', 'Giocatore', default=giocatore_scelto)
        ruolo = ottieni_valore('R', 'Ruolo', default='N.D.')
        squadra = ottieni_valore('Squadra', default='N.D.')

        st.subheader(f"👤 {nome_display}")
        st.write(f"**Ruolo:** {ruolo} | **Squadra:** {squadra}")
        st.divider()

        # Ricerca di Prezzi e Quotazioni potenziata
        p_max = ottieni_valore('Prezzo Max', 'Max', default=0)
        p_min = ottieni_valore('Prezzo Min', 'Min', default=0)
        p_med = ottieni_valore('Prezzo Medio', 'Medio', default=0)
        fvm = ottieni_valore('FVM', 'Consigliato', default='N.D.')

        # Gestione errori sicura: se le celle contengono testo o simboli, formatta a zero.
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
            st.write(f"**Presenze:** {ottieni_valore('Presenze', 'Pv', 'P', default=0)}")
            st.write(f"**Media Voto (MV):** {ottieni_valore('MV', 'Media Voto', default=0.0)}")
            st.write(f"**Fantamedia (FM):** {ottieni_valore('FM', 'Fantamedia', default=0.0)}")
            st.write(f"**Gol Fatti:** {ottieni_valore('Gol', 'Gf', default=0)} | **Assist:** {ottieni_valore('Assist', 'Ass', default=0)}")
            st.write(f"**Ammonizioni:** {ottieni_valore('Ammonizioni', 'Amm', 'Au', default=0)}")

        with col2:
            st.subheader('💰 Quotazioni 2026/2027')
            # Ricerca testuale infallibile
            qt_i = ottieni_valore('Qt.I', 'Iniziale', default='N.D.')
            qt_a = ottieni_valore('Qt.A', 'Attuale', default='N.D.')
            st.write(f"**Quotazione Iniziale (Qt.I):** {qt_i}")
            st.write(f"**Quotazione Attuale (Qt.A):** {qt_a}")
else:
    st.error('File "FANTA APP.xlsx" non trovato o impossibile leggere i nomi. Assicurati di aver fatto il Reboot dell\'app.')