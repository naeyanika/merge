import streamlit as st
import pandas as pd
import numpy as np
import io

st.title('Merge Pinjaman, TLP, dan KDP N/A')
st.write("""1. Ini digunakan untuk pivot table data dari pinjaman, tlp dan kdp N/A yang sudah di cari manual""")
st.write("""2. Nama data sesuai yang di download sebelumnya tidak usah di ubah lagi dan ektensi file nya dibiarkan .xlsx""")

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

        new_columns1 = [
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

        for col in new_columns1:
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
            'ID ANGGOTA', 'DUMMY', 'NAMA', 'CENTER', 'KEL', 'HARI', 'JAM', 'SL', 'TRANS. DATE',
            'Db PTN', 'Cr PTN', 'Db PRT', 'Cr PRT', 'Db DTP', 'Cr DTP', 'Db PMB', 'Cr PMB', 'Db PRR', 'Cr PRR',
            'Db PSA', 'Cr PSA', 'Db PU', 'Cr PU', 'Db Total2', 'Cr Total2'
        ]

        # Tambahkan kolom yang mungkin belum ada dalam DataFrame
        for col in desired_order:
            if col not in pivot_table1.columns:
                pivot_table1[col] = 0

        pivot_table1 = pivot_table1[desired_order]

        st.write("Pivot THC Pinjaman N/A:")
        st.write(pivot_table1)

    # Proses TLP N/A
    if 'TLP_na.xlsx' in dfs:
        df2 = dfs['TLP_na.xlsx']
        
        def sum_lists(x):
            if isinstance(x, list):
                return sum(int(value.replace('Rp ', '').replace(',', '')) for value in x)
            return x 

        df2['TRANS. DATE'] = pd.to_datetime(df2['TRANS. DATE'], format='%d/%m/%Y').dt.strftime('%d%m%Y')
        df2['DUMMY'] = df2['ID ANGGOTA'] + '' + df2['TRANS. DATE']

        pivot_table2 = pd.pivot_table(
            df2,
            values=['DEBIT', 'CREDIT'],
            index=['ID ANGGOTA', 'DUMMY', 'NAMA', 'CENTER', 'KELOMPOK', 'HARI', 'JAM', 'SL', 'TRANS. DATE'],
            columns='JENIS PINJAMAN',
            aggfunc={'DEBIT': list, 'CREDIT': list},
            fill_value=0
        )

        pivot_table2 = pivot_table2.applymap(sum_lists)
        pivot_table2.columns = [f'{col[0]}_{col[1]}' for col in pivot_table2.columns]
        pivot_table2.reset_index(inplace=True)
        pivot_table2['TRANS. DATE'] = pd.to_datetime(pivot_table2['TRANS. DATE'], format='%d%m%Y').dt.strftime('%d/%m/%Y')

        new_columns2 = [
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

        for col in new_columns2:
            if col not in pivot_table2.columns:
                pivot_table2[col] = 0

        pivot_table2['DEBIT_TOTAL'] = pivot_table2.filter(like='DEBIT').sum(axis=1)
        pivot_table2['CREDIT_TOTAL'] = pivot_table2.filter(like='CREDIT').sum(axis=1)

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

        pivot_table2 = pivot_table2.rename(columns=rename_dict)

        desired_order = [
            'ID ANGGOTA', 'DUMMY', 'NAMA', 'CENTER', 'KEL', 'HARI', 'JAM', 'SL', 'TRANS. DATE',
            'Db PTN', 'Cr PTN', 'Db PRT', 'Cr PRT', 'Db DTP', 'Cr DTP', 'Db PMB', 'Cr PMB', 'Db PRR', 'Cr PRR',
            'Db PSA', 'Cr PSA', 'Db PU', 'Cr PU', 'Db Total2', 'Cr Total2'
        ]

        # Tambahkan kolom yang mungkin belum ada dalam DataFrame
        for col in desired_order:
            if col not in pivot_table2.columns:
                pivot_table2[col] = 0

        pivot_table2 = pivot_table2[desired_order]

        st.write("Pivot TLP N/A:")
        st.write(pivot_table2)

    # Proses KDP N/A
    if 'KDP_na.xlsx' in dfs:
        df3 = dfs['KDP_na.xlsx']
        
        def sum_lists(x):
            if isinstance(x, list):
                return sum(int(value.replace('Rp ', '').replace(',', '')) for value in x)
            return x 

        df3['TRANS. DATE'] = pd.to_datetime(df3['TRANS. DATE'], format='%d/%m/%Y').dt.strftime('%d%m%Y')
        df3['DUMMY'] = df3['ID ANGGOTA'] + '' + df3['TRANS. DATE']

        pivot_table3 = pd.pivot_table(
            df3,
            values=['DEBIT', 'CREDIT'],
            index=['ID ANGGOTA', 'DUMMY', 'NAMA', 'CENTER', 'KELOMPOK', 'HARI', 'JAM', 'SL', 'TRANS. DATE'],
            columns='JENIS PINJAMAN',
            aggfunc={'DEBIT': list, 'CREDIT': list},
            fill_value=0
        )

        pivot_table3 = pivot_table3.applymap(sum_lists)
        pivot_table3.columns = [f'{col[0]}_{col[1]}' for col in pivot_table3.columns]
        pivot_table3.reset_index(inplace=True)
        pivot_table3['TRANS. DATE'] = pd.to_datetime(pivot_table3['TRANS. DATE'], format='%d%m%Y').dt.strftime('%d/%m/%Y')

        new_columns3 = [
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

        for col in new_columns3:
            if col not in pivot_table3.columns:
                pivot_table3[col] = 0

        pivot_table3['DEBIT_TOTAL'] = pivot_table3.filter(like='DEBIT').sum(axis=1)
        pivot_table3['CREDIT_TOTAL'] = pivot_table3.filter(like='CREDIT').sum(axis=1)

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

        pivot_table3 = pivot_table3.rename(columns=rename_dict)

        desired_order = [
            'ID ANGGOTA', 'DUMMY', 'NAMA', 'CENTER', 'KEL', 'HARI', 'JAM', 'SL', 'TRANS. DATE',
            'Db PTN', 'Cr PTN', 'Db PRT', 'Cr PRT', 'Db DTP', 'Cr DTP', 'Db PMB', 'Cr PMB', 'Db PRR', 'Cr PRR',
            'Db PSA', 'Cr PSA', 'Db PU', 'Cr PU', 'Db Total2', 'Cr Total2'
        ]

        # Tambahkan kolom yang mungkin belum ada dalam DataFrame
        for col in desired_order:
            if col not in pivot_table3.columns:
                pivot_table3[col] = 0

        pivot_table3 = pivot_table3[desired_order]

        st.write("Pivot KDP N/A:")
        st.write(pivot_table3)    

         # Download links for pivot tables
    for name, df in {
        'pivot_pinjaman_na.xlsx': pivot_table1,
        'pivot_TLP_na.xlsx': pivot_table2,
        'pivot_KDP_na.xlsx': pivot_table3
    }.items():
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
