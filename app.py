import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stats Fantacalcio", layout="centered")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@st.cache_data
def carica_dati():
  # Ricerca dinamica dei file basata su parole chiave (così non teme variazioni di nome o suffissi _2)
  file_2526, file_2627, file_tot, file_nuovi = None, None, None, None

  for f in os.listdir(BASE_DIR):
    if f.endswith('.xlsx'):
      if 'Statistiche_Fantacalcio_Stagione_2025_26' in f:
        file_2526 = os.path.join(BASE_DIR, f)
      elif 'Quotazioni_Fantacalcio_Stagione_2026_27' in f:
        file_2627 = os.path.join(BASE_DIR, f)
      elif f.startswith('Totale') or f == 'Totale.xlsx':
        file_tot = os.path.join(BASE_DIR, f)
      elif 'Nuovi_Acquisti' in f:
        file_nuovi = os.path.join(BASE_DIR, f)

  # Fallback di sicurezza sui file rilevati
  df_2526 = (
      pd.read_excel(file_2526, sheet_name='Tutti', header=1)
      if file_2526 and os.path.exists(file_2526)
      else pd.DataFrame()
  )
  df_2627 = (
      pd.read_excel(file_2627, sheet_name='Tutti', header=1)
      if file_2627 and os.path.exists(file_2627)
      else pd.DataFrame()
  )
  df_tot = (
      pd.read_excel(file_tot, sheet_name='Statistiche_Incrociate')
      if file_tot and os.path.exists(file_tot)
      else pd.DataFrame()
  )

  df_nuovi = pd.DataFrame()
  if file_nuovi and os.path.exists(file_nuovi):
    try:
      xls_n = pd.ExcelFile(file_nuovi)
      df_nuovi = pd.read_excel(file_nuovi, sheet_name=xls_n.sheet_names[0])
    except Exception:
      pass

  # Estrazione prezzi minimi dai fogli completi in Totale
  xls_tot = pd.ExcelFile(file_tot) if file_tot and os.path.exists(file_tot) else None
  complete_sheets = [
      'Portieri_Completo',
      'Difensori_Completo',
      'Centrocampo_Completo',
      'Attacco_Completo',
  ]
  all_role_sheets = []
  if xls_tot:
    for sheet in complete_sheets:
      if sheet in xls_tot.sheet_names:
        df_sheet = pd.read_excel(file_tot, sheet_name=sheet)
        header_indices = df_sheet[df_sheet.iloc[:, 0] == 'Nome Giocatore'].index
        for idx in header_indices:
          sub = df_sheet.iloc[idx + 1 :].copy()
          valid_rows = []
          for _, row in sub.iterrows():
            val0 = str(row.iloc[0])
            if (
                pd.isna(row.iloc[0])
                or 'TOP' in val0.upper()
                or 'MEDI' in val0.upper()
                or 'LOW' in val0.upper()
                or '🌟' in val0
            ):
              break
            valid_rows.append(row)
          if valid_rows:
            sub_clean = pd.DataFrame(valid_rows)
            sub_clean.columns = [
                'Nome Giocatore',
                'Frequenza',
                'Prezzo Max',
                'Prezzo Medio',
                'Prezzo Min',
            ]
            all_role_sheets.append(sub_clean)

  if all_role_sheets:
    df_prices = pd.concat(all_role_sheets, ignore_index=True)
    df_prices['Nome_Clean'] = (
        df_prices['Nome Giocatore'].astype(str).str.strip().str.lower()
    )
    df_prices_unique = (
        df_prices.groupby('Nome_Clean')['Prezzo Min'].min().reset_index()
    )
  else:
    df_prices_unique = pd.DataFrame(columns=['Nome_Clean', 'Prezzo Min'])

  # Pulizia chiavi testuali
  if not df_tot.empty and 'Nome Giocatore' in df_tot.columns:
    df_tot['Nome_Clean'] = (
        df_tot['Nome Giocatore'].astype(str).str.strip().str.lower()
    )
  else:
    df_tot['Nome_Clean'] = ''

  if not df_2627.empty and 'Nome' in df_2627.columns:
    df_2627['Nome_Clean'] = df_2627['Nome'].astype(str).str.strip().str.lower()
  else:
    df_2627['Nome_Clean'] = ''

  # Conversione ID in stringhe pulite per match sicuro
  if not df_2627.empty and 'Id' in df_2627.columns:
    df_2627['Id_Key'] = (
        pd.to_numeric(df_2627['Id'], errors='coerce')
        .dropna()
        .astype(int)
        .astype(str)
        .str.strip()
    )
  else:
    df_2627['Id_Key'] = ''

  if not df_2526.empty and 'Id' in df_2526.columns:
    df_2526['Id_Key'] = (
        pd.to_numeric(df_2526['Id'], errors='coerce')
        .dropna()
        .astype(int)
        .astype(str)
        .str.strip()
    )
  else:
    df_2526['Id_Key'] = ''

  # Merge principale
  merged = (
      pd.merge(
          df_tot,
          df_2627[['Nome_Clean', 'Id_Key', 'Qt.A', 'Qt.I']],
          on='Nome_Clean',
          how='left',
      )
      if not df_tot.empty
      else pd.DataFrame()
  )
  if not merged.empty and not df_prices_unique.empty:
    merged = pd.merge(merged, df_prices_unique, on='Nome_Clean', how='left')

  if not merged.empty and not df_2526.empty and 'Id_Key' in df_2526.columns:
    merged = pd.merge(
        merged,
        df_2526,
        left_on='Id_Key',
        right_on='Id_Key',
        how='left',
        suffixes=('', '_2526'),
    )

  # Integrazione nuovi acquisti tramite ID
  if not df_nuovi.empty and not merged.empty:
    id_col = [c for c in df_nuovi.columns if 'id' in str(c).lower()]
    if id_col:
      df_nuovi['Id_Key'] = (
          pd.to_numeric(df_nuovi[id_col[0]], errors='coerce')
          .dropna()
          .astype(int)
          .astype(str)
          .str.strip()
      )
      merged = pd.merge(
          merged,
          df_nuovi,
          on='Id_Key',
          how='left',
          suffixes=('', '_nuovi'),
      )

  # Valori di default numerici puliti
  for col in [
      'Presenze 25/26',
      'MV 25/26',
      'FM 25/26',
      'Gol 25/26',
      'Assist 25/26',
      'Ammoniz. 25/26',
  ]:
    if col not in merged.columns:
      merged[col] = 0
    else:
      merged[col] = merged[col].fillna(0)

  if 'Prezzo Medio' in merged.columns:
    merged['Prezzo Medio Int'] = (
        pd.to_numeric(merged['Prezzo Medio'], errors='coerce')
        .round()
        .fillna(0)
        .astype(int)
    )
  else:
    merged['Prezzo Medio Int'] = 0

  if 'Freq. Asta' in merged.columns:
    merged['Rank_Freq'] = (
        pd.to_numeric(merged['Freq. Asta'], errors='coerce')
        .rank(ascending=False, method='min')
        .fillna(999)
        .astype(int)
    )
  else:
    merged['Rank_Freq'] = 999

  return merged


df = carica_dati()

# --- INTERFACCIA UTENTE ---
st.title('⚽ Dashboard Fantacalcio & Asta')
st.markdown(
    'Cerca un giocatore per analizzare lo storico, i prezzi d’asta e le'
    ' statistiche aggiornate.'
)

if not df.empty and 'Nome Giocatore' in df.columns:
  lista_giocatori = sorted(df['Nome Giocatore'].dropna().unique().tolist())
  lista_giocatori.insert(0, '')

  giocatore_scelto = st.selectbox(
      'Scrivi o seleziona il nome del giocatore:', lista_giocatori
  )

  if giocatore_scelto:
    dati = df[df['Nome Giocatore'] == giocatore_scelto].iloc[0]

    st.subheader(f"👤 {dati.get('Nome Giocatore', giocatore_scelto)}")
    ruolo = dati.get('Ruolo', dati.get('R', 'N.D.'))
    squadra = dati.get('Squadra 26/27', dati.get('Squadra', 'N.D.'))
    st.write(f'**Ruolo:** {ruolo} | **Squadra 26/27:** {squadra}')
    st.divider()

    col_a1, col_a2, col_a3, col_a4 = st.columns(4)
    col_a1.metric('Prezzo Max', f"{dati.get('Prezzo Max', 0)} cr.")
    col_a2.metric(
        'Prezzo Min',
        f"{int(dati.get('Prezzo Min', 0))} cr."
        if pd.notna(dati.get('Prezzo Min'))
        else 'N.D.',
    )
    col_a3.metric('Prezzo Medio', f"{dati.get('Prezzo Medio Int', 0)} cr.")
    col_a4.metric('FVM Consigliato', f"{dati.get('FVM 26/27', 'N.D.')}")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
      st.subheader('📊 Statistiche / Storico')
      st.write(f"**Presenze:** {dati.get('Presenze 25/26', 0)}")
      st.write(f"**Media Voto (MV):** {dati.get('MV 25/26', 0.0)}")
      st.write(f"**Fantamedia (FM):** {dati.get('FM 25/26', 0.0)}")
      st.write(
          f"**Gol Fatti:** {dati.get('Gol 25/26', 0)} | **Assist:**"
          f" {dati.get('Assist 25/26', 0)}"
      )
    with col2:
      st.subheader('💰 Quotazioni 2026/2027')
      st.write(f"**Quotazione Iniziale (Qt.I):** {dati.get('Qt.I', 'N.D.')}")
      st.write(f"**Quotazione Attuale (Qt.A):** {dati.get('Qt.A', 'N.D.')}")
else:
  st.error(
      'Errore nel caricamento dei dati. Assicurati che i file Excel siano'
      ' presenti nella cartella.'
  )