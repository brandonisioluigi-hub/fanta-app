import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stats Fantacalcio", layout="centered")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@st.cache_data
def carica_dati():
  file_2526, file_2627, file_tot = None, None, None

  for f in os.listdir(BASE_DIR):
    if f.endswith('.xlsx') and not f.startswith('~$'):
      f_lower = f.lower()
      if 'statistiche' in f_lower:
        file_2526 = f
      elif 'quotazioni' in f_lower:
        file_2627 = f
      elif 'totale' in f_lower:
        file_tot = f

  # Caricamento file principali
  df_tot_incrociate = (
      pd.read_excel(
          os.path.join(BASE_DIR, file_tot), sheet_name='Statistiche_Incrociate'
      )
      if file_tot
      else pd.DataFrame()
  )
  df_2627 = (
      pd.read_excel(
          os.path.join(BASE_DIR, file_2627), sheet_name='Tutti', header=1
      )
      if file_2627
      else pd.DataFrame()
  )
  df_2526 = (
      pd.read_excel(
          os.path.join(BASE_DIR, file_2526), sheet_name='Tutti', header=1
      )
      if file_2526
      else pd.DataFrame()
  )

  # Estrazione prezzi da TUTTI i fogli completi del file Totale.xlsx
  df_prices_all = pd.DataFrame(
      columns=['Nome_Clean', 'Prezzo Max', 'Prezzo Medio', 'Prezzo Min']
  )
  if file_tot:
    try:
      xls_tot = pd.ExcelFile(os.path.join(BASE_DIR, file_tot))
      list_subs = []
      for sheet in xls_tot.sheet_names:
        if 'completo' in sheet.lower() or 'statistiche' in sheet.lower():
          df_s = pd.read_excel(xls_tot, sheet_name=sheet)
          header_indices = df_s[
              df_s.astype(str)
              .apply(lambda col: col.str.contains('Nome Giocatore', na=False))
              .any(axis=1)
          ].index
          for idx in header_indices:
            sub = df_s.iloc[idx + 1 :].copy()
            sub = sub.dropna(subset=[sub.columns[0]])
            valid_rows = []
            for _, row in sub.iterrows():
              val0 = str(row.iloc[0])
              if (
                  'TOP' in val0.upper()
                  or 'MEDI' in val0.upper()
                  or 'LOW' in val0.upper()
                  or '🌟' in val0
                  or val0 == 'nan'
              ):
                break
              valid_rows.append(row)
            if valid_rows:
              sub_clean = pd.DataFrame(valid_rows)
              if sub_clean.shape[1] >= 5:
                sub_clean = sub_clean.iloc[:, :5]
                sub_clean.columns = [
                    'Nome Giocatore',
                    'Frequenza',
                    'Prezzo Max',
                    'Prezzo Medio',
                    'Prezzo Min',
                ]
                list_subs.append(sub_clean)

      if list_subs:
        dp = pd.concat(list_subs, ignore_index=True)
        dp['Nome_Clean'] = dp['Nome Giocatore'].astype(str).str.strip().str.lower()
        dp['Prezzo Max'] = pd.to_numeric(dp['Prezzo Max'], errors='coerce')
        dp['Prezzo Medio'] = pd.to_numeric(dp['Prezzo Medio'], errors='coerce')
        dp['Prezzo Min'] = pd.to_numeric(dp['Prezzo Min'], errors='coerce')
        df_prices_all = (
            dp.groupby('Nome_Clean')[['Prezzo Max', 'Prezzo Medio', 'Prezzo Min']]
            .first()
            .reset_index()
        )
    except Exception:
      pass

  # Uniformiamo chiavi e ID
  if not df_tot_incrociate.empty:
    df_tot_incrociate['Nome_Clean'] = (
        df_tot_incrociate['Nome Giocatore']
        .astype(str)
        .str.strip()
        .str.lower()
    )
  if not df_2627.empty:
    df_2627['Nome_Clean'] = df_2627['Nome'].astype(str).str.strip().str.lower()
    df_2627['Id'] = pd.to_numeric(df_2627['Id'], errors='coerce')
  if not df_2526.empty:
    df_2526['Id'] = pd.to_numeric(df_2526['Id'], errors='coerce')

  # Unione
  merged = df_2627.copy()
  if 'Nome' in merged.columns and 'Nome Giocatore' not in merged.columns:
    merged['Nome Giocatore'] = merged['Nome']

  if not df_2526.empty and 'Id' in df_2526.columns:
    merged = pd.merge(merged, df_2526, on='Id', how='left', suffixes=('', '_2526'))

  if not df_tot_incrociate.empty:
    merged = pd.merge(
        merged,
        df_tot_incrociate,
        on='Nome_Clean',
        how='left',
        suffixes=('', '_tot'),
    )

  if not df_prices_all.empty:
    merged = pd.merge(merged, df_prices_all, on='Nome_Clean', how='left')

  return merged


df = carica_dati()

# --- INTERFACCIA UTENTE ---
st.title('⚽ Dashboard Fantacalcio & Asta')

if not df.empty and 'Nome Giocatore' in df.columns:
  lista_giocatori = sorted(df['Nome Giocatore'].dropna().unique().tolist())
  lista_giocatori.insert(0, '')

  giocatore_scelto = st.selectbox(
      'Scrivi o seleziona il nome del giocatore:', lista_giocatori
  )

  if giocatore_scelto:
    dati = df[df['Nome Giocatore'] == giocatore_scelto].iloc[0]


    def get_val(col_name):
      if col_name in dati and pd.notna(dati[col_name]):
        return dati[col_name]
      return 0


    st.subheader(f"👤 {dati.get('Nome Giocatore', giocatore_scelto)}")
    ruolo = dati.get('R', dati.get('Ruolo', 'N.D.'))
    squadra = dati.get('Squadra', 'N.D.')
    st.write(f'**Ruolo:** {ruolo} | **Squadra 26/27:** {squadra}')
    st.divider()

    p_max = dati.get('Prezzo Max')
    p_min = dati.get('Prezzo Min')
    p_med = dati.get('Prezzo Medio')

    col_a1, col_a2, col_a3, col_a4 = st.columns(4)
    col_a1.metric(
        'Prezzo Max', f'{int(p_max)} cr.' if pd.notna(p_max) else '0 cr.'
    )
    col_a2.metric(
        'Prezzo Min', f'{int(p_min)} cr.' if pd.notna(p_min) else 'N.D.'
    )
    col_a3.metric(
        'Prezzo Medio', f'{int(p_med)} cr.' if pd.notna(p_med) else '0 cr.'
    )
    col_a4.metric(
        'FVM Consigliato', f"{dati.get('FVM', dati.get('FVM 26/27', 'N.D.'))}"
    )

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
      st.subheader('📊 Statistiche / Storico')
      st.write(
          f"**Presenze:** {get_val('Pv') if 'Pv' in dati else get_val('Presenze')}"
      )
      st.write(
          f"**Media Voto (MV):** {get_val('Mv') if 'Mv' in dati else get_val('MV')}"
      )
      st.write(
          f"**Fantamedia (FM):** {get_val('Fm') if 'Fm' in dati else get_val('FM')}"
      )
      st.write(
          f"**Gol Fatti:** {get_val('Gf') if 'Gf' in dati else get_val('Gol')} |"
          f" **Assist:** {get_val('Ass') if 'Ass' in dati else get_val('Assist')}"
      )
      st.write(
          f"**Ammonizioni:** {get_val('Amm') if 'Amm' in dati else get_val('Ammonizioni')}"
      )

    with col2:
      st.subheader('💰 Quotazioni 2026/2027')
      # Variabili estratte pulite per evitare qualsiasi SyntaxError nelle f-string
      qt_i = dati.get('Qt.I', 'N.D.')
      qt_a = dati.get('Qt.A', 'N.D.')
      st.write(f'**Quotazione Iniziale (Qt.I):** {qt_i}')
      st.write(f'**Quotazione Attuale (Qt.A):** {qt_a}')
else:
  st.error('Errore nel caricamento dei dati. Verifica la presenza dei file.')