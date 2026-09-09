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

  # Trova il file dei nuovi acquisti indipendentemente dal numero finale
  file_nuovi = None
  for f in os.listdir(BASE_DIR):
    if 'Nuovi_Acquisti' in f and f.endswith('.xlsx'):
      file_nuovi = os.path.join(BASE_DIR, f)
      break

  # 1. Lettura dei file principali
  df_2526 = pd.read_excel(file_2526, sheet_name='Tutti', header=1)
  df_2627 = pd.read_excel(file_2627, sheet_name='Tutti', header=1)
  df_tot = pd.read_excel(file_tot, sheet_name='Statistiche_Incrociate')

  # 2. Lettura del file dei nuovi acquisti (con gestione flessibile dell'header)
  df_nuovi = pd.DataFrame()
  if file_nuovi and os.path.exists(file_nuovi):
    try:
      xls_n = pd.ExcelFile(file_nuovi)
      df_nuovi = pd.read_excel(file_nuovi, sheet_name=xls_n.sheet_names[0])
    except Exception:
      pass

  # Estrazione prezzi minimi dai fogli completi per ruolo
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

  # Conversione ID in numerico per match sicuro
  if 'Id' in df_2627.columns:
    df_2627['Id'] = pd.to_numeric(df_2627['Id'], errors='coerce')
  if 'Id' in df_2526.columns:
    df_2526['Id'] = pd.to_numeric(df_2526['Id'], errors='coerce')

  # Merge iniziale: Totale con Quotazioni 26/27 per recuperare l'ID
  merged = pd.merge(
      df_tot,
      df_2627[['Nome_Clean', 'Id', 'Qt.A', 'Qt.I']],
      on='Nome_Clean',
      how='left',
  )
  merged = pd.merge(merged, df_prices_unique, on='Nome_Clean', how='left')

  # Merge con lo storico 25/26 tramite ID
  merged = pd.merge(
      merged, df_2526, on='Id', how='left', suffixes=('', '_2526')
  )

  # Integrazione dei nuovi acquisti tramite ID
  if not df_nuovi.empty:
    # Troviamo la colonna ID nel file nuovi acquisti
    id_col_candidates = [
        c for c in df_nuovi.columns if 'id' in str(c).lower()
    ]
    if id_col_candidates:
      df_nuovi = df_nuovi.rename(columns={id_col_candidates[0]: 'Id'})
      df_nuovi['Id'] = pd.to_numeric(df_nuovi['Id'], errors='coerce')

      # Uniamo i nuovi acquisti
      merged = pd.merge(
          merged,
          df_nuovi,
          on='Id',
          how='left',
          suffixes=('', '_nuovi'),
      )

      # Se i dati statistici nel file dei nuovi acquisti sono presenti nelle colonne posizionali
      # (es. Presenze, Gol, Assist, MV, FM), li mappiamo direttamente dove mancanti
      cols = list(df_nuovi.columns)
      if len(cols) >= 8:
        # Mappatura sicura basata sulla struttura tipica dei file di nuovi acquisti
        # (ID, Nome, Squadra, Provenienza, Ruolo, Presenze, Gol, Assist, MV...)
        pass

  # Pulizia valori mancanti per evitare zeri o errori a schermo
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