import pandas as pd
import streamlit as st
from sentence_transformers import SentenceTransformer , util
from rapidfuzz import process, fuzz
from collections import Counter
import io

def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> io.BytesIO:
    """
    Convierte un DataFrame a un archivo Excel en memoria y devuelve un BytesIO listo
    para usar en st.download_button.
    - No escribe a disco.
    - Mantiene tipos (número/texto) tal como estén en el DataFrame.
    - sheet_name: nombre de la hoja dentro del .xlsx
    """
    buffer = io.BytesIO()
    # podés usar engine="xlsxwriter" u "openpyxl"; ambos funcionan
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        # acá podrías agregar formatos, anchos de columna, etc.
    buffer.seek(0)  # muy importante: rebobinar para que Streamlit lea desde el inicio
    return buffer

st.title("Autoimputador de Cuentas")

try:
    Saldo_Inicio = st.text_input("Ingrese el saldo inicial del periodo (Con coma o punto):")
    if Saldo_Inicio == "":
        Saldo_Inicio = 0
except TypeError:
    st.stop()

if Saldo_Inicio != 0:
    try:
        Saldo_Inicio = float(Saldo_Inicio.replace(",","."))
    except ValueError:
        st.warning("Escriba un numero valido porfavor.")
        st.stop()

try:
    Saldo_Final = st.text_input("Ingrese el saldo final del periodo:")
    if Saldo_Final == "":
        Saldo_Final = 0 
except TypeError:
    st.stop()

if Saldo_Final != 0:
    try:
        Saldo_Final = float(Saldo_Final.replace(",","."))
    except ValueError:
        st.warning("Escriba un numero valido porfavor.")
        st.stop()

Variacion_Saldo = Saldo_Final - Saldo_Inicio
Variacion_Saldo = round(Variacion_Saldo, 2)

def coma_a_punto(x):
    return str(x).replace(",",".")

with st.container(border=True):
    st.subheader("Sube Aqui los General Ledgers de este año y el anterior en formato CSV")
    archivos = st.file_uploader("sube aqui los GLS", accept_multiple_files=True, type=["csv"])

lista_de_GLS = []
if archivos:  
    for f in archivos:
        df_temp = pd.read_csv(f, encoding="latin1")
        if not df_temp.empty:
            lista_de_GLS.append(df_temp)

if not lista_de_GLS:
    st.info("Subí al menos un CSV válido para continuar.")
    st.stop()

df_original = pd.concat(lista_de_GLS, ignore_index=True)

#creacion del df a trabajar
df_a_trabajar = pd.DataFrame()
df_a_trabajar["Name"] = df_original["Name"]
df_a_trabajar["Memo"] = df_original["Memo"]
df_a_trabajar["Split"] = df_original["Split"]

#Armado de diccionario
namelist= []
filas, columnas = df_a_trabajar.shape

x = 0
while x < filas:
    name = ""
    name = df_a_trabajar["Name"].iloc[x]
    x+=1
    if pd.isna(name):
        continue
    else:
        namelist.append(name)
namelist = list(set(namelist))

with st.container(border=True):
    st.subheader("Suba Aqui el Accountlist Actual en formato Excel")
    account_archive = st.file_uploader("Sube Aqui el account list.", type=["xlsx"])

if not account_archive:
    st.info("Subí el Accountlist (XLSX) para continuar.")
    st.stop()

accounts_df = pd.read_excel(account_archive, sheet_name="Sheet1")

accounts_df_mejorado = pd.DataFrame(columns=["Accounts Principales", "Accounts Secundarias"])

accountlist_banks = []
accountlist_without_banks = []
accountlist_banks = accounts_df[accounts_df["Type"]=="Bank"]["Account"].tolist()
accountlist_without_banks = accounts_df[accounts_df["Type"]!="Bank"]["Account"].tolist()


#Armado de la cuenta principal y la secundaria
x = 0
for i in accountlist_without_banks:
    i = str(i)
    finder = i.find(":")
    if finder > 0:
        accounts_df_mejorado.at[x, "Accounts Principales"] = i
        accounts_df_mejorado.at[x, "Accounts Secundarias"] = i[finder+1:]
        x+=1
    else:
        accounts_df_mejorado.at[x, "Accounts Principales"] = i 
        x+= 1


#Armado del excel de vendors con sus splits
memos = []
split = []
filtrados = []
df_name_filtro = pd.DataFrame(columns=["Name", "Memos", "Split"])
df_name_filtro["Name"] = namelist
df_name_filtro = df_name_filtro.set_index("Name")
for i in namelist: 
    memos.extend(df_a_trabajar[df_a_trabajar["Name"]==i]["Memo"].tolist())  
    df_name_filtro.at[i,"Memos"] = list(set(memos))
    memos = []
    split.extend(df_a_trabajar[df_a_trabajar["Name"]==i]["Split"].tolist())  
    for j in split:
        similitud, bestscore, index = process.extractOne(query=j, scorer= fuzz.partial_ratio, choices= accountlist_banks)
        similitud2, bestscore2, index2 = process.extractOne(query=j, scorer= fuzz.WRatio, choices=accounts_df_mejorado["Accounts Principales"].tolist())
        similitud3, bestscore3, index3 = process.extractOne(query=j, scorer= fuzz.WRatio, choices=accounts_df_mejorado["Accounts Secundarias"].tolist())
        if bestscore > 80:
            continue
        elif bestscore2==100:
            filtrados.append(similitud2)
        elif bestscore3>80:
            filtrados.append(accounts_df_mejorado["Accounts Principales"].iloc[index3])
        else:
            filtrados.append(j)
    conteo = Counter(filtrados)
    if conteo:
        conteo = conteo.most_common(1)[0][0]
    else:
        conteo = None
    df_name_filtro.at[i,"Split"] = conteo
    split = []
    filtrados = []

#Armado del excel de memos con sus splits
df_memo_filtro = pd.DataFrame()
df_memo_filtro["Memo"] = df_original[df_original["Memo"].notna()]["Memo"]
df_memo_filtro["Split"] = df_original[df_original["Memo"].notna()]["Split"]

lista_de_splits = []
lista_de_memos = []

df_memo_filtro = df_memo_filtro.reset_index(drop = True)
x = 0
for i in df_memo_filtro["Split"]:
    similitud, bestscore, index = process.extractOne(query=i, scorer= fuzz.partial_ratio, choices= accountlist_banks)
    if bestscore > 80:
        x+=1
        continue
    else:
        similitud2, bestscore2 , index2 = process.extractOne(query=i, scorer= fuzz.WRatio, choices= accounts_df_mejorado["Accounts Principales"].tolist())
        similitud3, bestscore3 , index3 = process.extractOne(query=i, scorer= fuzz.WRatio, choices= accounts_df_mejorado["Accounts Secundarias"].tolist())
        if bestscore2 == 100:
            lista_de_splits.append(similitud2)
        elif bestscore3 > 80:
            lista_de_splits.append(accounts_df_mejorado["Accounts Principales"].iloc[index3])
        else:
            lista_de_splits.append(i)
        lista_de_memos.append(df_memo_filtro["Memo"].iloc[x])
        x+= 1
df_nuevo_memo = pd.DataFrame()
df_nuevo_memo["Memo"]= lista_de_memos
df_nuevo_memo["Split"] = lista_de_splits


#Carga de Archivos
with st.container(border=True):
    st.subheader("Suba Aqui los Archivos generados por los Modelos en Formato CSV, UNA CUENTA A LA VEZ.")
    CSVs_archives = st.file_uploader("Sube aqui los archivos.", accept_multiple_files=True, type=["csv"])

lista_de_archivos = [] 
if CSVs_archives:  
    for i in CSVs_archives:
        df_temp = pd.read_csv(i, encoding="latin1")
        if not df_temp.empty:  # Evitar agregar DataFrames vacíos
            lista_de_archivos.append(df_temp)

if not lista_de_archivos:
    st.info("Subí al menos un CSV de modelos válido para continuar.")
    st.stop()

df_archivos_subidos = pd.concat(lista_de_archivos, ignore_index=True)

Sumatoria_archivos_subidos = round(df_archivos_subidos["Amount"].sum(), 2)

if Sumatoria_archivos_subidos - Variacion_Saldo != 0:
    st.warning(f"La Variacion de los Movimientos debe ser igual a la diferencia entre el Saldo Final y el Saldo Inicial, la diferencia es de {Sumatoria_archivos_subidos - Variacion_Saldo}, la sumatoria de los movimientos de los archivos subidos son {Sumatoria_archivos_subidos}")
    st.stop()


#reseteo del index de df_name_filtro:
df_name_filtro.reset_index(inplace=True)

#usar IA con tensores
model = SentenceTransformer("all-MiniLM-L6-v2")
emb_accounts = model.encode(accounts_df["Account"].tolist(), convert_to_tensor=True)

#Armado de los deposits
df_Deposits = pd.DataFrame(columns= ["Date", "Vendor", "Account", "Description", "Check", "Amount ABS", "Revisar"])
df_Deposits["Date"] = df_archivos_subidos[df_archivos_subidos["Amount"]>0]["Date"]
df_Deposits["Description"] = df_archivos_subidos[df_archivos_subidos["Amount"]>0]["Description"]
df_Deposits["Amount ABS"] = df_archivos_subidos[df_archivos_subidos["Amount"]>0]["Amount"]
df_Deposits ["Amount ABS"] = df_Deposits["Amount ABS"].apply(coma_a_punto)
if "Check" in df_archivos_subidos.columns:
    df_Deposits["Check"] = df_archivos_subidos[df_archivos_subidos["Amount"]>0]["Check"]
df_Deposits.reset_index(inplace=True)
df_Deposits.drop(columns="index", inplace=True)

x= 0
for i in df_Deposits["Description"]:
    lower = str(i).lower()
    similitud1, bestscore1, index1 = process.extractOne(query= lower, scorer= fuzz.partial_ratio, choices= df_name_filtro["Name"].str.lower().tolist())
    similitud2, bestscore2, index2 = process.extractOne(query= i, scorer= fuzz.WRatio, choices= df_nuevo_memo["Memo"].tolist())
    if bestscore1 > 80:
        df_Deposits.at[x, "Vendor"] = df_name_filtro["Name"].iloc[index1]
        df_Deposits.at[x, "Account"] = df_name_filtro["Split"].iloc[index1]
        x+=1
    elif bestscore2 > 80:
        df_Deposits.at[x, "Account"] = df_nuevo_memo[df_nuevo_memo["Memo"] == similitud2]["Split"].iloc[0]
        Split_ = df_nuevo_memo[df_nuevo_memo["Memo"] == similitud2]["Split"].iloc[0]
        if Split_ == "-SPLIT-":
            df_Deposits.at[x, "Revisar"] = "Revisar"
        x+=1
    else:
            emb_memo = model.encode(i, convert_to_tensor=True)
            Escores_de_coseno = util.cos_sim(emb_memo,emb_accounts)[0]
            Index_del_mejorcoseno = Escores_de_coseno.argmax().item()
            df_Deposits.at[x, "Account"] = accounts_df["Account"].iloc[Index_del_mejorcoseno]
            df_Deposits.at[x, "Revisar"] = "Revisar"
            x+=1


if "df_Deposits" in locals() and not df_Deposits.empty:
    deposits_xlsx = df_to_excel_bytes(df_Deposits, sheet_name="Deposits")
    st.download_button(
        label="Descargar Deposits.xlsx",
        data=deposits_xlsx,
        file_name="Deposits.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

#Armado de Checks
def niunamenos (x):
    return x.replace("-","")

df_Checks = pd.DataFrame(columns= ["Date", "Check", "Vendor", "Account", "Amount ABS", "Description", "Revisar"])
df_Checks["Date"] = df_archivos_subidos[df_archivos_subidos["Amount"]<0]["Date"]
df_Checks["Description"] = df_archivos_subidos[df_archivos_subidos["Amount"]<0]["Description"]
df_Checks["Amount ABS"] = df_archivos_subidos[df_archivos_subidos["Amount"]<0]["Amount"]
df_Checks["Amount ABS"] = df_Checks["Amount ABS"].apply(coma_a_punto)
df_Checks["Amount ABS"] = df_Checks["Amount ABS"].apply(niunamenos)
if "Check" in df_archivos_subidos.columns:
    df_Checks["Check"] = df_archivos_subidos[df_archivos_subidos["Amount"]<0]["Check"]
df_Checks.reset_index(inplace=True)
df_Checks.drop(columns="index", inplace=True)

x= 0
for i in df_Checks["Description"]:
    lower = str(i).lower()
    similitud1, bestscore1, index1 = process.extractOne(query= lower, scorer= fuzz.partial_ratio, choices= df_name_filtro["Name"].str.lower().tolist())
    similitud2, bestscore2, index2 = process.extractOne(query= i, scorer= fuzz.WRatio, choices= df_nuevo_memo["Memo"].tolist())
    if bestscore1 > 80:
        df_Checks.at[x, "Vendor"] = df_name_filtro["Name"].iloc[index1]
        df_Checks.at[x, "Account"] = df_name_filtro["Split"].iloc[index1]
        x+=1
    elif bestscore2 > 80:
        df_Checks.at[x, "Account"] = df_nuevo_memo[df_nuevo_memo["Memo"] == similitud2]["Split"].iloc[0]
        Split_ = df_nuevo_memo[df_nuevo_memo["Memo"] == similitud2]["Split"].iloc[0]
        if Split_ == "-SPLIT-":
            df_Checks.at[x, "Revisar"] = "Revisar"
        x+=1
    else:
        emb_memo = model.encode(i, convert_to_tensor=True)
        Escores_de_coseno = util.cos_sim(emb_memo,emb_accounts)[0]
        Index_del_mejorcoseno = Escores_de_coseno.argmax().item()
        df_Checks.at[x, "Account"] = accounts_df["Account"].iloc[Index_del_mejorcoseno]
        df_Checks.at[x, "Revisar"] = "Revisar"
        x+=1


if "df_Checks" in locals() and not df_Checks.empty:
    checks_xlsx = df_to_excel_bytes(df_Checks, sheet_name="Checks")
    st.download_button(
        label="Descargar Checks.xlsx",
        data=checks_xlsx,
        file_name="Checks.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

#Armado de Credit Cards
df_Creditcard = pd.DataFrame(columns= ["Date", "Vendor", "Account", "Amount ABS", "Description", "Revisar"])
df_Creditcard["Date"] = df_archivos_subidos["Date"]
df_Creditcard["Description"] = df_archivos_subidos["Description"]
df_Creditcard["Amount ABS"] = df_archivos_subidos["Amount"]
df_Creditcard["Amount ABS"] = df_Creditcard["Amount ABS"].apply(coma_a_punto)
df_Creditcard.reset_index(inplace=True)
df_Creditcard.drop(columns="index", inplace=True)

x= 0
for i in df_Creditcard["Description"]:
    lower = str(i).lower()
    similitud1, bestscore1, index1 = process.extractOne(query= lower, scorer= fuzz.partial_ratio, choices= df_name_filtro["Name"].str.lower().tolist())
    similitud2, bestscore2, index2 = process.extractOne(query= i, scorer= fuzz.WRatio, choices= df_nuevo_memo["Memo"].tolist())
    if bestscore1 > 80:
        df_Creditcard.at[x, "Vendor"] = df_name_filtro["Name"].iloc[index1]
        df_Creditcard.at[x, "Account"] = df_name_filtro["Split"].iloc[index1]
        x+=1
    elif bestscore2 > 80:
        df_Creditcard.at[x, "Account"] = df_nuevo_memo[df_nuevo_memo["Memo"] == similitud2]["Split"].iloc[0]
        Split_ = df_nuevo_memo[df_nuevo_memo["Memo"] == similitud2]["Split"].iloc[0]
        if Split_ == "-SPLIT-":
            df_Creditcard.at[x, "Revisar"] = "Revisar"
        x+=1
    else:
        emb_memo = model.encode(i, convert_to_tensor=True)
        Escores_de_coseno = util.cos_sim(emb_memo,emb_accounts)[0]
        Index_del_mejorcoseno = Escores_de_coseno.argmax().item()
        df_Creditcard.at[x, "Account"] = accounts_df["Account"].iloc[Index_del_mejorcoseno]
        df_Creditcard.at[x, "Revisar"] = "Revisar"
        x+=1

if "df_Creditcard" in locals() and not df_Creditcard.empty:
    credit_xlsx = df_to_excel_bytes(df_Creditcard, sheet_name="CreditCard")
    st.download_button(
        label="Descargar CreditCard.xlsx",
        data=credit_xlsx,
        file_name="CreditCard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )