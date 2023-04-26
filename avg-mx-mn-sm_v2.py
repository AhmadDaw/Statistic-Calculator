from tkinter import *
import pandas as pd
from tkinter import filedialog
import customtkinter as ctk
import numpy as np

ctk.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("green")  # Themes: "blue" (standard), "green", "dark-blue"

app = ctk.CTk()
app.geometry("310x480")
app.title("Statistic Calculator v2")

root = ctk.CTkFrame(master=app)
root.pack(pady=20, padx=20, fill="both", expand=True)

e_avg=ctk.CTkEntry(root, width=120)
e_mx=ctk.CTkEntry(root, width=120)
e_mn=ctk.CTkEntry(root, width=120)
e_sm=ctk.CTkEntry(root, width=120)

l_avg=ctk.CTkLabel(root, text="Average: ")
l_mx=ctk.CTkLabel(root, text="Max: ")
l_mn=ctk.CTkLabel(root, text="Min: ")
l_sm=ctk.CTkLabel(root, text="Sum: ")

l=ctk.CTkLabel(root, text=" ")
ll=ctk.CTkLabel(root, text=" ")
stats=ctk.CTkLabel(root, text="-")
stats_cln=ctk.CTkLabel(root, text="-")

kml_title=ctk.CTkLabel(root, text="Statistic Calculator")
data_cln_title=ctk.CTkLabel(root, text="Statistic Calculator")

#status = Label(root, text = "Coded by: Ahmad Dawara", bd=2, relief=SUNKEN, anchor = E)

# -----------------------------------------------------
def chk():
    return
# ----------------------------------------------------
var1 = IntVar()
vr1 = ctk.CTkCheckBox(root, text='Extract Cell Name.',variable=var1, onvalue=1, offvalue=0, command=chk)

def cln():
    print(fp[-3:])
    if fp[-3:] =='csv':
        df = pd.read_csv(fp)
    elif (fp[-4:] =='xlsx') or (fp[-3:] =='xls'):
        df = pd.read_excel(fp)
    else:
        stats_cln.configure(text='Please Choose csv or Excel file.')

    if (var1.get() == 1):
        col_name=df.columns[3]
        #print(col_name)
        if (col_name=='Cell'):
            df['a'] = df[str(col_name)].str.slice(start=21)
            print(df.head())
            df[['r','x', 'y','z']] = df['a'].str.split(',',3, expand=True)
            print(df.head())
            df[['b','Cell']] = df['y'].str.split('=',1, expand=True)
            print(df.head())
            col_x_name=df.columns[2]
            df.rename(columns={str(col_x_name): 'Site'}, inplace=True)
            df = df.drop(['r','b','a','x','y','z'], axis=1, errors='ignore')
            col_x_name=df.columns[2]
            print(df.head())
        else:
            if 'TRX' in df.columns:
                df.drop(df['TRX'], axis=1,inplace=True, errors='ignore')
            df.rename(columns={str(col_name): 'Cells'}, inplace=True)
            col_name=df.columns[3]
            df['a'] = df[str(col_name)].str.slice(start=6)
            df[['Cell','b','c']] = df['a'].str.split(',',2, expand=True)
            df[str(col_name)]=df['Cell']
            df[['Site','x']] = df[str(col_name)].str.split('_',1, expand=True)
            scol_name=df.columns[2]
            df[str(scol_name)]=df['Site']
            df = df.drop(['b','a','c','x','Site','Cell'], axis=1, errors='ignore')
            df.rename(columns={str(scol_name): 'Site'}, inplace=True)
    
    # ----------------------------------------------------------------
    s_avg='a'
    s_mx='a'
    s_mn='a'
    s_sm='a'
    
    s_avg=e_avg.get()
    s_mx=e_mx.get()
    s_mn=e_mn.get()
    s_sm=e_sm.get()

    def col_num_to_names_lst(lstx):
        if (';' in lstx):
            a_list=lstx.split(';')
            n_list=[]

            for cn in a_list:
                cn=int(cn)
                cn=cn-1
                if cn in range(len(df.columns)):
                    n_list.append(df.columns[int(cn)])
            #print(n_list)
            return(n_list)
        elif (';' not in lstx) and lstx!='':
            if int(lstx) in range(len(df.columns)):
                n_list=[]
                x=int(lstx)
                x=x-1
                n_list.append(df.columns[int(x)])
                return(n_list)            
        else:
            return None

    n_avg_list=col_num_to_names_lst(s_avg)
    n_mx_list=col_num_to_names_lst(s_mx)
    n_mn_list=col_num_to_names_lst(s_mn)
    n_sm_list=col_num_to_names_lst(s_sm)

    print(n_avg_list)
    print(n_mx_list)
    print(n_mn_list)
    print(n_sm_list)
    
    df['Start Time']=df['Start Time'].astype(str)
    print(df)
    df[['Dates','Times']]= df['Start Time'].str.split(' ',1, expand=True)
    #df.iloc[:,4:]  =  df.iloc[:,4:].apply(pd.to_numeric)
    df.iloc[:,4:] = df.iloc[:,4:].replace('NIL', np.nan, regex=True)

    excel_file = 'cell-kpi-agg.xlsx'
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

    if n_avg_list!=None:
        d_avg = {n_avg_list[i]: ['mean'] for i in range(len(n_avg_list))}
        
        df_avg = df.groupby(['Dates', 'Cells']).agg(d_avg)
        df_avg = df_avg.reset_index()
        df_avg.to_excel(writer, sheet_name='Avg')

    if n_mx_list!= None:
        d_mx={n_mx_list[i]: ['max'] for i in range(len(n_mx_list))}
        df_mx = df.groupby(['Dates', 'Cells']).agg(d_mx)
        df_mx = df_mx.reset_index()
        df_mx.to_excel(writer, sheet_name='Max')

    if n_sm_list!= None:
        d_sm={n_sm_list[i]: ['sum'] for i in range(len(n_sm_list))}
        df_sm = df.groupby(['Dates', 'Cells']).agg(d_sm)
        df_sm = df_sm.reset_index()
        df_sm.to_excel(writer, sheet_name='Sum')

    if n_mn_list!= None:
        d_mn={n_mn_list[i]: ['min'] for i in range(len(n_mn_list))}
        df_mn = df.groupby(['Dates', 'Cells']).agg(d_mn)
        df_mn = df_mn.reset_index()
        df_mn.to_excel(writer, sheet_name='Min')

    df.to_csv('cleaned-kpis.csv',index=False)
    writer.save()  
    stats_cln.configure(text='Done')
# -------------------------------------------------------------------
def opn_cln():
    global fp
    fp=filedialog.askopenfilename()
# -------------------------------------------------------------
b_opn_cln=ctk.CTkButton(root, text="Browse" ,command=opn_cln)
b_browse_cln=ctk.CTkButton(root, text="Clean File" ,command=cln)

# ----------------------------------------------------------
data_cln_title.grid(row = 1, column = 2,pady=10,padx=10)
b_opn_cln.grid(row = 2, column = 2, pady=10)
# ---------------------------------------------------------
e_avg.grid(row = 3, column = 2, pady=10)
e_mx.grid(row = 4, column = 2, pady=10)
e_mn.grid(row = 5, column = 2, pady=10)
e_sm.grid(row = 6, column = 2, pady=10)

l_avg.grid(row = 3, column = 1, pady=10)
l_mx.grid(row = 4, column = 1, pady=10)
l_mn.grid(row = 5, column = 1, pady=10)
l_sm.grid(row = 6, column = 1, pady=10)
# ---------------------------------------------------------
vr1.grid(row = 7, column = 2, pady=10)
stats_cln.grid(row = 8, column = 2, pady=10, padx=10)
b_browse_cln.grid(row = 9, column = 2, pady=10)
# ---------------------------------------------------------
l.grid(row = 1, column = 1, pady=10, padx=10)
ll.grid(row = 1, column = 0, pady=10, padx=10)
# ----------------------------------------------------------
app.mainloop()


