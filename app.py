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

        def ottieni_valore(*nomi_colonne, default=0):
            for col in nomi_colonne:
                col_cercata = str(col).replace(" ", "").replace(".", "").lower()
                
                # Match esatto ripulito
                match_col = next((c for c in dati.index if str(c).replace(" ", "").replace(".", "").lower() == col_cercata), None)
                
                # Fallback: cerca anche solo una parola (es. "Iniziale" in "Quotazione Iniziale")
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

        # Estrazione Valori Flessibile
        fvm = ottieni_valore('FVM', 'Consigliato', 'FVM 26/27', default='N.D.')
        qt_i = ottieni_valore('Qt.I', 'Iniziale', 'Qt Iniziale', 'Quotazione', default='N.D.')
        qt_a = ottieni_valore('Qt.A', 'Attuale', 'Qt Attuale', default='N.D.')
        
        p_max = ottieni_valore('Prezzo Max', 'Max', 'Prezzo Massimo', default=0)
        p_min = ottieni_valore('Prezzo Min', 'Min', 'Prezzo Minimo', default=0)
        p_med = ottieni_valore('Prezzo Medio', 'Medio', default=0)

        # Helper per trasformare in numero
        def pulisci_numero(val):
            try:
                return float(str(val).replace(',', '.'))
            except:
                return 0.0

        val_qta = pulisci_numero(qt_a)
        val_qti = pulisci_numero(qt_i)
        
        # Paracadute: se il prezzo d'asta non c'è, usa la quotazione.
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