import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stats Fantacalcio", layout="centered")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@st.cache_data
def carica_dati():
  file_2526 = os.path.join(
      BASE_DIR, 'Statistiche_Fantacalcio_Stagione_2025_26.xlsx'
  )
  file_2627 = os.path.join(
      BASE_DIR, 'Quotazioni_Fantacalcio_Stagione_2026_27 (1).xlsx'
  )
  file_tot = os.path.join(BASE_DIR, 'Totale.xlsx')

  # Cerchiamo dinamicamente il file dei nuovi acquisti indipendentemente dal numerino finale
  file_nuovi = None
  for f in os.listdir(BASE_DIR):
    if 'Nuovi_Acquisti' in f and f.endswith('.xlsx'):
      file_nuovi = os.path.join(BASE_DIR, f)
      break

  # Caricamento dei file principali
  df_2526 = pd.read_excel(file_2526, sheet_name='Tutti', header=1)
  df_2627 = pd.read_excel(file_2627, sheet_name='Tutti', header=1)
  df_tot = pd.read_excel(file_tot, sheet_name='Statistiche_Incrociate')

  # Caricamento nuovi acquisti se esiste
  df_nuovi = pd.DataFrame()
  if file_nuovi and os.path.exists(file_nuovi):
    xls_n = pd.ExcelFile(file_nuovi)
    df_nuovi = pd.read_excel(file_nuovi, sheet_name=xls_n.sheet_names[0])

  # Estrazione prezzi minimi dai fogli completi per ruolo in Totale.xlsx
  xls_tot = pd.ExcelFile(file_tot)
  complete_sheets = [
      'Portieri_Completo',
      'Difensori_Completo',
      'Centrocampo_Completo',
      'Attacco_Completo',
  ]
  all_role_sheets = []
  for sheet in complete_sheets:
    if sheet in xls_tot.sheet_names:
      df_sheet = pd.read_excel(xls_tot, sheet_name=sheet)
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

  # Pulizia e uniformazione chiavi di collegamento
  df_tot['Nome_Clean'] = (
      df_tot['Nome Giocatore'].astype(str).str.strip().str.lower()
  )
  df_2627['Nome_Clean'] = df_2627['Nome'].astype(str).str.strip().str.lower()

  # 1. Uniamo Totale con le Quotazioni 26/27 tramite il nome pulito per recuperare l'ID
  merged = pd.merge(
      df_tot,
      df_2627[['Nome_Clean', 'Id', 'Qt.A', 'Qt.I']],
      on='Nome_Clean',
      how='left',
  )
  merged = pd.merge(merged, df_prices_unique, on='Nome_Clean', how='left')

  # 2. Unione con lo storico 25/26 tramite ID
  merged = pd.merge(
      merged, df_2526, on='Id', how='left', suffixes=('', '_2526')
  )

  # 3. Se abbiamo il file dei nuovi acquisti, integriamo i dati mancanti tramite ID
  if not df_nuovi.empty and 'Id' in df_nuovi.columns:
    merged = pd.merge(
        merged, df_nuovi, on='Id', how='left', suffixes=('', '_nuovi')
    )
    # Se alcune colonne chiave arrivano dai nuovi acquisti, ripuliamo i valori nulli
    for col in [
        'Presenze',
        'MV',
        'FM',
        'Gol',
        'Assist',
        'Ammoniz',
        'Presenze 25/26',
        'MV 25/26',
        'FM 25/26',
        'Gol 25/26',
        'Assist 25/26',
    ]:
      col_n = f'{col}_nuovi'
      if col_n in merged.columns and col in merged.columns:
        merged[col] = merged[col].fillna(merged[col_n])
      elif col_n in merged.columns:
        merged[col] = merged[n] if (col not in merged.columns) else merged[col]

  # Formattazioni finali
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


@st.cache_data
def carica_ranking(sheet_name):
  file_tot = os.path.join(BASE_DIR, 'Totale.xlsx')
  xls = pd.ExcelFile(file_tot)
  if sheet_name not in xls.sheet_names:
    return {}
  df = pd.read_excel(xls, sheet_name=sheet_name)
  rank_dict = {}
  current_role = None
  rank_in_role = 0

  for _, row in df.iterrows():
    val0 = str(row.iloc[0])
    if 'AFFIDABILITÀ' in val0.upper() or 'PRESTAZIONI' in val0.upper():
      if 'POR' in val0.upper():
        current_role = 'POR'
      elif 'DIF' in val0.upper():
        current_role = 'DIF'
      elif 'CEN' in val0.upper():
        current_role = 'CEN'
      elif 'ATT' in val0.upper():
        current_role = 'ATT'
      rank_in_role = 0
      continue
    if val0 == 'Nome Giocatore' or pd.isna(row.iloc[1]):
      continue
    rank_in_role += 1
    rank_dict[val0.strip().lower()] = (current_role, rank_in_role)
  return rank_dict


df = carica_dati()
aff_dict = carica_ranking('Top_Affidabilita')
pres_dict = carica_ranking('Top_Prestazioni')

# --- INTERFACCIA UTENTE ---
st.title('⚽ Dashboard Fantacalcio & Asta')
st.markdown(
    'Cerca un giocatore per analizzare lo storico, i prezzi d’asta e le'
    ' statistiche aggiornate.'
)

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

  st.info('📋 **Dati Simulazioni & Prezzi Asta**')
  col_a1, col_a2, col_a3, col_a4, col_a5 = st.columns(5)

  freq_val = int(dati['Freq. Asta']) if pd.notna(dati.get('Freq. Asta')) else 0
  rank_val = int(dati['Rank_Freq']) if pd.notna(dati.get('Rank_Freq')) else 0
  col_a1.metric('Frequenza Asta', f'{freq_val} volte', f'{rank_val}° in classifica')
  col_a2.metric('Prezzo Max', f"{dati.get('Prezzo Max', 0)} cr.")

  p_min = dati.get('Prezzo Min')
  p_min_str = f'{int(p_min)} cr.' if pd.notna(p_min) else 'N.D.'
  col_a3.metric('Prezzo Min', p_min_str)

  col_a4.metric('Prezzo Medio', f"{dati.get('Prezzo Medio Int', 0)} cr.")
  col_a5.metric('FVM Consigliato', f"{dati.get('FVM 26/27', 'N.D.')}")

  st.divider()

  col1, col2 = st.columns(2)

  with col1:
    st.subheader('📊 Statistiche / Storico')
    presenze = dati.get(
        'Presenze 25/26', dati.get('Presenze', dati.get('P', '0'))
    )
    mv = dati.get('MV 25/26', dati.get('MV', '0.00'))
    fm = dati.get('FM 25/26', dati.get('FM', '0.00'))
    gol = dati.get('Gol 25/26', dati.get('Gol', dati.get('Gf', '0')))
    assist = dati.get('Assist 25/26', dati.get('Assist', '0'))
    ammoniz = dati.get('Ammoniz. 25/26', dati.get('Ammoniz.', '0'))

    st.write(f'**Presenze:** {presenze}')
    st.write(f'**Media Voto (MV):** {mv}')
    st.write(f'**Fantamedia (FM):** {fm}')
    st.write(f'**Gol Fatti:** {gol} | **Assist:** {assist}')

    if str(ruolo).strip() == 'POR':
      st.write(f"**Gol Subiti:** {dati.get('Gs', 0)}")

    st.write(f'**Ammonizioni:** {ammoniz}')

    nome_clean = str(dati.get('Nome Giocatore', '')).strip().lower()
    aff = aff_dict.get(nome_clean)
    if aff:
      st.success(f'⭐ **Top Affidabilità:** {aff[1]}° tra i {aff[0]}')

    pres = pres_dict.get(nome_clean)
    if pres:
      st.success(f'🔥 **Top Prestazioni:** {pres[1]}° tra i {pres[0]}')

  with col2:
    st.subheader('💰 Quotazioni 2026/2027')
    qt_a = dati.get('Qt.A')
    qt_i = dati.get('Qt.I')
    st.write(
        f"**Quotazione Iniziale (Qt.I):**"
        f" {qt_i if pd.notna(qt_i) else 'N.D.'}"
    )
    st.write(
        f"**Quotazione Attuale (Qt.A):**"
        f" {qt_a if pd.notna(qt_a) else 'N.D.'}"
    )
    st.write(f"**Diff. Max Asta:** {dati.get('Diff. Max', 'N.D.')}")
    diff_medio = dati.get('Diff. Medio')
    diff_medio_str = f'{diff_medio:.2f}' if pd.notna(diff_medio) else 'N.D.'
    st.write(f"**Diff. Medio Asta:** {diff_medio_str}")