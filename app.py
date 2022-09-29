import streamlit as st
import pandas as pd
import unidecode
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import numpy as np
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import streamlit as st
from stqdm import stqdm
from PIL import Image



def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1',encoding='latin1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data


st.title('Outil de matching IEG')

st.subheader('Inserez un ficher excel qui respecte la norme convenu')
#st.markdown("<h1 style='text-align: center; color: red;'>Inserez un ficher excel qui respecte la norme convenu</h1>", unsafe_allow_html=True)
st.markdown("""
* Feuille 1 
  * Veuillez mettre la liste des éléments à matcher dans la première colonne de la première feuille et mettez le titre ou laisser la première cellule vide
* Feuille 2
  * Veuillez mettre la liste source dans la première colonne de la deuxième feuille et mettez le titre ou laisser la première cellule vide""")


uploaded_file = st.file_uploader("Choose an excel file",type=['xlsx','xls'], accept_multiple_files=False,label_visibility="hidden")



if uploaded_file is not None:
    

    with st.spinner("Veuillez attendre s'il vous plaît..."):
        df1 = pd.read_excel(uploaded_file,sheet_name=0)
        df2 = pd.read_excel(uploaded_file,sheet_name=1)




        df1['col_trait']= df1[df1.columns[0]]
        df2['col_trait'] = df2[df2.columns[0]]

        df1 = df1.replace({np.nan:None})
        df2 = df2.replace({np.nan:None})



        df1['col_trait'] = df1['col_trait'].apply(lambda x : x.lower() if x else None)
        df1['col_trait'] = df1['col_trait'].apply(lambda x :unidecode.unidecode(x) if x else None)
        df1['col_trait'] = df1['col_trait'].apply(lambda x :x.replace(".","") if x else None)
        df1['col_trait'] = df1['col_trait'].apply(lambda x :x.replace(",","") if x else None)
        df1['col_trait'] = df1['col_trait'].apply(lambda x :x.replace(" ","") if x else None)
        df1['col_trait'] = df1['col_trait'].apply(lambda x :x.replace("'","") if x else None)
        df1['col_trait'] = df1['col_trait'].apply(lambda x :x.replace('"',"") if x else None)
        df1['col_trait'] = df1['col_trait'].apply(lambda x :x.replace('ep',"") if x else None)
        df1['col_trait'] = df1['col_trait'].apply(lambda x :x.replace('epouse',"") if x else None)
        df1['col_trait'] = df1['col_trait'].apply(lambda x :x.replace('epous',"") if x else None)





        df2['col_trait'] = df2['col_trait'].apply(lambda x : x.lower() if x else None)
        df2['col_trait'] = df2['col_trait'].apply(lambda x :unidecode.unidecode(x) if x else None)
        df2['col_trait'] = df2['col_trait'].apply(lambda x :x.replace(".","") if x else None)
        df2['col_trait'] = df2['col_trait'].apply(lambda x :x.replace(",","") if x else None)
        df2['col_trait'] = df2['col_trait'].apply(lambda x :x.replace(" ","") if x else None)
        df2['col_trait'] = df2['col_trait'].apply(lambda x :x.replace("'","") if x else None)
        df2['col_trait'] = df2['col_trait'].apply(lambda x :x.replace('"',"") if x else None)
        df2['col_trait'] = df2['col_trait'].apply(lambda x :x.replace('epous',"") if x else None)
        df2['col_trait'] = df2['col_trait'].apply(lambda x :x.replace('ep',"") if x else None)
        df2['col_trait'] = df2['col_trait'].apply(lambda x :x.replace('epouse',"") if x else None)


        df2 = df2.drop_duplicates(subset=['col_trait'],keep='first')

        dic_df2 = {}
        for y,x in zip(df2[df2.columns[0]],df2['col_trait']):
            dic_df2[x]=y

        L = []
        choices = list(dic_df2.keys())

        for x in stqdm(df1['col_trait']):

            L.append(process.extract(x, choices, limit=2))


        df1["Choix 1"] = [dic_df2[x[0][0]] for x in L]
        df1['Score choix 1'] = [x[0][1] for x in L]

        df1["Choix 2"] = [dic_df2[x[1][0]] for x in L]
        df1['Score choix 2'] = [x[1][1] for x in L]

        df1 = df1.drop(columns={"col_trait"})
        
        if len(df2.columns != 1):
            df1 = df1.merge(df2, how='left',left_on="Choix 1", right_on=df2.columns[0])

        st.dataframe(df1)
        df_xlsx = to_excel(df1)

    st.download_button(label='📥 Télécharger le résultat',
                                data=df_xlsx ,
                                file_name= 'resultats_matching.xlsx')


