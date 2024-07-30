import streamlit as st
import pandas as pd
import numpy as np
import io

st.title('Merge Simpanan, Pinjaman, dan N/A Pinjaman')

uploaded_files = st.file_uploader("Unggah file Excel", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    dfs = {}
    for file in uploaded_files:
        df = pd.read_excel(file, engine='openpyxl')  # Baca file Excel dengan pandas
        dfs[file.name] = df
    
    # Proses Pinjaman N/A
    if 'pinjaman_na.xlsx' in dfs:
        df1 = dfs['pinjaman_na.xlsx']
        
        def sum_lists(x):
            if isinstance(x, list):
                return sum(int(value.replace('Rp ', '').replace(',', '')) for value in x)
            return x 

        df1['TRANS. DATE'] = pd.to_datetime(df1['TRANS. DATE'], format='%d/%m/%Y').dt.strftime('%d%m%Y')
        df1['DUMMY'] = df1['ID ANGGOTA'] + '' + df1['TRANS. DATE']

        pivot_table1 = pd.pivot_table(
            df1,
            values=['DEBIT', 'CREDIT'],
            index=['ID ANGGOTA', 'DUMMY', 'NAMA', 'CENTER', 'KELOMPOK', 'HARI', 'JAM', 'SL', 'TRANS. DATE'],
            columns='JENIS PINJAMAN',
            aggfunc={'DEBIT': list, 'CREDIT': list},
            fill_value=0
        )

        pivot_table1 = pivot_table1.applymap(sum_lists)
        pivot_table1.columns = [f'{col[0]}_{col[1]}' for col in pivot_table1.columns]
        pivot_table1.reset_index(inplace=True)
        pivot_table1['TRANS. DATE'] = pd.to_datetime(pivot_table1['TRANS. DATE'], format='%d%m%Y').dt.strftime('%d/%m/%Y')

        new_columns4 = [
            'DEBIT_PINJAMAN UMUM',
            'DEBIT_PINJAMAN RENOVASI RUMAH',
            'DEBIT_PINJAMAN SANITASI',
            'DEBIT_PINJAMAN ARTA',
            'DEBIT_PINJAMAN MIKROBISNIS',
            'DEBIT_PINJAMAN DT. PENDIDIKAN',
            'DEBIT_PINJAMAN PERTANIAN',
            'CREDIT_PINJAMAN UMUM',
            'CREDIT_PINJAMAN RENOVASI RUMAH',
            'CREDIT_PINJAMAN SANITASI',
            'CREDIT_PINJAMAN ARTA',
            'CREDIT_PINJAMAN MIKROBISNIS',
            'CREDIT_PINJAMAN DT. PENDIDIKAN',
            'CREDIT_PINJAMAN PERTANIAN'
        ]

        for col in new_columns4:
            if col not in pivot_table1.columns:
                pivot_table1[col] = 0

        pivot_table1['DEBIT_TOTAL'] = pivot_table1.filter(like='DEBIT').sum(axis=1)
        pivot_table1['CREDIT_TOTAL'] = pivot_table1.filter(like='CREDIT').sum(axis=1)

        rename_dict = {
            'KELOMPOK': 'KEL',
            'DEBIT_PINJAMAN ARTA': 'Db PRT',
            'DEBIT_PINJAMAN DT. PENDIDIKAN': 'Db DTP',
            'DEBIT_PINJAMAN MIKROBISNIS': 'Db PMB',
            'DEBIT_PINJAMAN SANITASI': 'Db PSA',
            'DEBIT_PINJAMAN UMUM': 'Db PU',
            'DEBIT_PINJAMAN RENOVASI RUMAH': 'Db PRR',
            'DEBIT_PINJAMAN PERTANIAN': 'Db PTN',
            'DEBIT_TOTAL': 'Db Total2',
            'CREDIT_PINJAMAN ARTA': 'Cr PRT',
            'CREDIT_PINJAMAN DT. PENDIDIKAN': 'Cr DTP',
            'CREDIT_PINJAMAN MIKROBISNIS': 'Cr PMB',
            'CREDIT_PINJAMAN SANITASI': 'Cr PSA',
            'CREDIT_PINJAMAN UMUM': 'Cr PU',
            'CREDIT_PINJAMAN RENOVASI RUMAH': 'Cr PRR',
            'CREDIT_PINJAMAN PERTANIAN': 'Cr PTN',
            'CREDIT_TOTAL': 'Cr Total2'
        }

        pivot_table1 = pivot_table1.rename(columns=rename_dict)

        desired_order = [
        'Db PTN', 'Cr PTN', 'Db ARTA', 'Cr ARTA', 'Db DTP', 'Cr DTP', 'Db PMB', 'Cr PMB', 
        'Db PRR', 'Cr PRR', 'Db PSA', 'Cr PSA', 'Db PU', 'Cr PU', 'Db Total2', 'Cr Total2'
        ]

         for col in desired_order:
            if col not in pivot_table1.columns:
                pivot_table1[col] = 0
        
        pivot_table1 = pivot_table1[desired_order]
        
        st.write("Pivot THC Pinjaman N/A:")
        st.write(pivot_table1)

        # Download links for pivot tables
        for name, df in {'pivot_pinjaman_na.xlsx': pivot_table1}.items():
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            buffer.seek(0)
            st.download_button(
                label=f"Unduh {name}",
                data=buffer.getvalue(),
                file_name=name,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
