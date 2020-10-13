# -*- coding: utf-8 -*-
"""
Created on Fri Jun  8 10:30:09 2018

@author: Marco Di Luzio
"""

from tkinter import *
import os
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askdirectory
import datetime
from tkinter import ttk
from tkinter import messagebox
try:
    import xlrd
    import numpy as np
    import matplotlib
    import xlsxwriter
    import pandas as pd
except ModuleNotFoundError:
    answ = input('Additional python modules have not been found.\nThe required modules (modules name and version are found in the requirements.txt file) will be installed in your current python environment, otherwise you will need to manually manage the additional packages.\nInternet connection is necessary.\n\nPress <y> key to confirm the modules installation or any other key to exit: ')
    if answ.lower()=='y':
        os.system('cmd /k "pip install -r requirements.txt"')
    quit()
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
try:
    from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk as NavigationToolbar2TkAgg
except:
    from matplotlib.backends.backend_tkagg import NavigationToolbar2TkAgg
from matplotlib.figure import Figure
from classes.rnaaobj import *

#version
VERSION = 1.1
VERSION_DATE = '8 October 2020'

#Matplotlib graph visualization parameters
matplotlib.rcParams['font.size'] = 8
matplotlib.rcParams['savefig.dpi'] = 300
matplotlib.rcParams['axes.formatter.limits'] = [-5,5]

def main():
    def initialization(settingsfile):
        f=open(settingsfile, 'r')
        rs=f.readlines()
        f.close()
        for i in range(len(rs)):
            rs[i]=rs[i].replace('\n','')
        return rs
    
    def findoutexceltable(wb,sheet):
        fs=wb.sheet_by_name(sheet)
        table=[]
        a=0
        while a>-1:
            try:
                vlx=str(fs.cell(0,a).value)
            except:
                break
            else:
                a=a+1
        table=[]
        r=1
        while r>0:
            try:
                line=[]
                for i in range(a):
                    vlx=str(fs.cell(r,i).value)
                    try:
                        line.append(float(vlx))
                    except:
                        line.append(vlx)
            except:
                break
            else:
                table.append(line)
                r=r+1
        return table
    
    def totalk0tables(wb,sssh):
        tablesk0databases=[]
        for index in sssh:
            tablesk0databases.append(findoutexceltable(wb,index))
        return tablesk0databases
    
    def openselectedk0database(k0database):
        """Open k0 database from file xls"""
        try:
            wb=xlrd.open_workbook('data/rel_data/'+k0database)
        except FileNotFoundError:
            pass
        else:
            sssh=wb.sheet_names()
            A=totalk0tables(wb,sssh)
            return A
        
    def lineacumulata(A):
        """Arrange k0 database on memory"""
        keyline=[]
        for i in range(len(A[0])-1):
            key=A[0][i][0]
            emission=A[0][i][:18]
            for decay in range(len(A[4])):
                if A[4][decay][0]==key:
                    emiss,fath,grandfa=A[4][decay][0],A[4][decay][5],A[4][decay][6]
                    decaycode=A[4][decay][:8]
                    break
            for nuclide in range(len(A[3])-1):
                if A[3][nuclide][0]==emiss:
                    daught=A[3][nuclide][:13]
                    break
            for nuclide in range(len(A[3])-1):
                if A[3][nuclide][0]==fath:
                    father=A[3][nuclide][:13]
                    break
                else:
                    father=['','','','','','','','','','','','','']
            for nuclide in range(len(A[3])-1):
                if A[3][nuclide][0]==grandfa:
                    grandfather=A[3][nuclide][:13]
                    break
                else:
                    grandfather=['','','','','','','','','','','','','']
            newkey=daught[0]
            capture=[]
            for capt in range(len(A[1])-2):
                if A[1][capt][14]==newkey:
                    capture=A[1][capt][:27]
                    break
            if capture==[]:
                newkey=father[0]
                for capt in range(len(A[1])-2):
                    if A[1][capt][14]==newkey:
                        capture=A[1][capt][:27]
                        break
            if capture==[]:
                newkey=grandfather[0]
                for capt in range(len(A[1])-2):
                    if A[1][capt][14]==newkey:
                        capture=A[1][capt][:27]
                        break
            if grandfather[0]!='':
                keyline.append(emission+decaycode+daught+grandfather+father+capture)
            else:
                keyline.append(emission+decaycode+daught+father+grandfather+capture)
        return keyline
    
    def emission_database(filename,sheets):
        """Open k0 database from file xls"""
        try:
            wb=xlrd.open_workbook('data/rel_data/'+filename)
        except FileNotFoundError:
            pass
        else:
            A = []
            for sheet in sheets:
                fs=wb.sheet_by_name(sheet)
                current_sheet = [fs.row_values(i,0,4)+[sheet] for i in range(1,fs.nrows-1)]
                A += current_sheet
            return A
    
    def td(S1,S2):
        TD=S1.datetime-S2.datetime
        return abs(TD.days*86400+TD.seconds)
    
    def listeffy_short(n):
        h=os.listdir('data/efficiencies')
        H=[]
        extensions=['.eny','.fwh','.eff']
        for t in h:
            if t[-4:]==extensions[n]:
                H.append(t[:-4])
        return H
    
    def delete_list(EC,NAA):
        extensions=['.eny','.fwh','.eff','.drr']
        if EC.get()!='':
            for k in extensions:
                os.remove('data/efficiencies/'+EC.get()+k)
            EC.set('')
            values=listeffy()
            EC['values']=values
            NAA.enegycomparatorfit=None
            NAA.enegysamplefit=None
            NAA.fwhmcomparatorfit=None
            NAA.fwhmsamplefit=None
            NAA.efficiencycomparatorfit=None
            NAA.dercomparatorfit=[None,None]
            NAA.efficiencysamplefit=None
            NAA.dersamplefit=[None,None]
    
    def recharge_lists(EC):
        E=listeffy_short(2)
        EC.configure(values=E)
        
    def rename_itemlist(EC,BT,NAA):
        def confirm_rename(E,EC,TLRN,NAA):
            listdir=os.listdir('data/efficiencies')
            if E.get()!='' and E.get()!=EC.get() and E.get()+'.eff' not in listdir:
                extensions=['.eny','.fwh','.eff','.drr']
                for ex in extensions:
                    os.replace('data/efficiencies/'+EC.get()+ex,'data/efficiencies/'+E.get()+ex)
                recharge_lists(EC)
                EC.set(E.get())
                NAA.efficiencycomparatorfit.name=E.get()
                NAA.efficiencysamplefit.name=E.get()
                TLRN.destroy()
            elif E.get()=='':
                print('An empty new calibration name is entered')
            elif E.get()==EC.get():
                print('The new calibration name should be different from the old one')
            elif E.get()+'.eff' in listdir:
                print('The new calibration name entered already exists')
        def return_confirm_rename(ew,E,EC,TLRN,NAA):
            confirm_rename(E,EC,TLRN,NAA)
        if EC.get()!='':
            h,w,a,b=EC.winfo_height(),EC.winfo_width(),EC.winfo_rootx(),EC.winfo_rooty()
            TLRN=Toplevel()
            TLRN.geometry(f'{w}x{h}+{a}+{b+h}')
            TLRN.overrideredirect(True)
            TLRN.resizable(False,False)
            E=Entry(TLRN)
            E.pack(side=LEFT, fill=X, expand=True)
            E.insert(0,EC.get())
            E.focus()            
            BC=Button(TLRN, text='Confirm', width=8, command=lambda E=E,EC=EC,TLRN=TLRN,NAA=NAA : confirm_rename(E,EC,TLRN,NAA)).pack(side=RIGHT)
            ew='<Return>'
            E.bind(ew,lambda ew=ew,E=E,EC=EC,TLRN=TLRN,NAA=NAA : return_confirm_rename(ew,E,EC,TLRN,NAA))
            event="<FocusOut>"
            TLRN.bind(event,lambda event=event,M=TLRN : M.destroy())
    
    def listeffy(f=None):
        h=os.listdir('data/efficiencies')
        H=[]
        n=2
        extensions=['.eny','.fwh','.eff']
        for t in h:
            if t[-4:]==extensions[n]:
                H.append(t[:-4])
        return H
    
    def openeffy(w):
        f=open(w,'r')
        r=f.readlines()
        f.close()
        R=[]
        for i in range(len(r)):
            R.append(str.split(r[i].replace('\n','')))
        return R
    
    def uncertainty_shaper(L,u):
        def selection(Rdint,RS1,RS2,RS3,ND,HD,TD):
            if Rdint.get()==1:
                RS1.select()
                RS2.deselect()
                RS3.deselect()
                ND.configure(state='normal')
                HD.configure(state='readonly')
                TD.configure(state='readonly')
            elif Rdint.get()==2:
                RS1.deselect()
                RS2.select()
                RS3.deselect()
                ND.configure(state='readonly')
                HD.configure(state='normal')
                TD.configure(state='readonly')
            else:
                RS1.deselect()
                RS2.deselect()
                RS3.select()
                ND.configure(state='readonly')
                HD.configure(state='readonly')
                TD.configure(state='normal')
        def effect(u,ND,HD,TD,TUS):
            DD=[None,ND,HD,TD]
            try:
                U=float(DD[int(Rdint.get())].get())
                if Rdint.get()==2:
                    U=np.sqrt(np.power(U,2)/3)
                elif Rdint.get()==3:
                    U=np.sqrt(np.power(U,2)/6)
            except:
                pass
            else:
                u.delete(0,END)
                u.insert(0,U)
                TUS.destroy()
        try:
            U=float(u.get())
        except:
            U=0.0
        TUS=Toplevel()
        TUS.title(L)
        TUS.resizable(False,False)
        L=Label(TUS).pack()
        Rdint=IntVar(TUS)
        F=Frame(TUS)
        L=Label(F, width=1).pack(side=LEFT)
        RS1=Radiobutton(F, text='Gaussian distribution - Standard uncertainty', anchor=W, variable=Rdint, value=1, width=40)
        RS1.pack(side=LEFT)
        ND=Spinbox(F, from_=0, to=1, increment=0.001, width=10)
        ND.delete(0,END)
        ND.insert(0,U)
        ND.pack(side=LEFT)
        F.pack(anchor=W)
        F=Frame(TUS)
        L=Label(F, width=1).pack(side=LEFT)
        RS2=Radiobutton(F, text='Rectangular distribution - Half-width interval', anchor=W, variable=Rdint, value=2, width=40)
        RS2.pack(side=LEFT)
        HD=Spinbox(F, from_=0, to=1, increment=0.001, width=10)
        HD.pack(side=LEFT)
        F.pack(anchor=W)
        F=Frame(TUS)
        L=Label(F, width=1).pack(side=LEFT)
        RS3=Radiobutton(F, text='Triangular distribution - Half-width interval', anchor=W, variable=Rdint, value=3, width=40)
        RS3.pack(side=LEFT)
        TD=Spinbox(F, from_=0, to=1, increment=0.001, width=10)
        TD.pack(side=LEFT)
        L=Label(F, width=1).pack(side=LEFT)
        F.pack(anchor=W)
        RS1.configure(command=lambda Rdint=Rdint,RS1=RS1,RS2=RS2,RS3=RS3,ND=ND,HD=HD,TD=TD : selection(Rdint,RS1,RS2,RS3,ND,HD,TD))
        RS2.configure(command=lambda Rdint=Rdint,RS1=RS1,RS2=RS2,RS3=RS3,ND=ND,HD=HD,TD=TD : selection(Rdint,RS1,RS2,RS3,ND,HD,TD))
        RS3.configure(command=lambda Rdint=Rdint,RS1=RS1,RS2=RS2,RS3=RS3,ND=ND,HD=HD,TD=TD : selection(Rdint,RS1,RS2,RS3,ND,HD,TD))
        TUS.focus()
        RS1.invoke()
        L=Label(TUS).pack()
        BAPY=Button(TUS, text='Apply', width=8, command= lambda u=u,ND=ND,HD=HD,TD=TD,TUS=TUS: effect(u,ND,HD,TD,TUS))
        BAPY.pack()
        L=Label(TUS).pack()
        
    def funcEnergy(x, a, b):
        return a*x+b
    
    def funcFWHM(x, a, b):
        return np.sqrt(a*x+b)
    
    def write_hyperlabfiles(wbname,TLHPL):
        wbname=wbname.cget('text')
        foldername=askdirectory()
        if foldername!='' and foldername!=None and wbname!='':
            wb = xlrd.open_workbook(wbname)
            ss=wb.sheet_names()
            head=['SERIAL#',"PEAK_ID","EMPLOYEE_ID","PEAKEVAL_ID","POS","POSUNC","E","EUNC","AREA","AREAUNC","FWHM","FWHMUNC","CREATIONDATE","PEAKEVALPEAK_ID","HEIGHT","HEIGHTUNC","BACKGROUND","EFF","EFFUNC","EFF.CORR.AREA","EFF.CORR.AREA UNC"]
            try:
                for sht in ss:
                    fs=wb.sheet_by_name(sht)
                    date=xlrd.xldate_as_tuple(fs.cell(0,1).value,0)#tuple(year,month,day,hour,minute,second)
                    real_t=float(fs.cell(1,1).value)
                    live_t=float(fs.cell(2,1).value)
                    n_chs=int(fs.cell(3,1).value)
                    RS=[]
                    xch=[]
                    xE=[]
                    xF=[]
                    r=5
                    while r>0:
                        try:
                            CH=float(fs.cell(r,0).value)
                            EG=float(fs.cell(r,1).value)
                            AR=float(fs.cell(r,2).value)
                            UA=float(fs.cell(r,3).value)
                            FW=fs.cell(r,4).value
                            if FW=='':
                                FW=0
                        except:
                            break
                        else:
                            line=['0','0','0','0',CH,'0',EG,'0',AR,UA,FW,'0','0','0','0','0','0','0','0','0','0']
                            RS.append(line)
                            xch.append(CH)
                            xE.append(EG)
                            xF.append(FW)
                            r+=1
                    Cch=[]
                    r=1
                    while r<n_chs+5:
                        try:
                            Cch.append(int(fs.cell(r,8).value))
                        except:
                            Cch.append(0)
                        r+=1
                    with open(foldername+'/'+sht+'.csv', 'w', newline='') as csvfile:
                        spamwriter = csv.writer(csvfile)
                        spamwriter.writerow(head)
                        spamwriter.writerows(RS)
                    with open(foldername+'/'+sht+'.ASC', 'w') as ASCfile:
                        for cn in range(n_chs):
                            ASCfile.write(f'{Cch[cn]}\n')
                        ASCfile.write(f'#LiveTime={live_t}\n')
                        ASCfile.write(f'#TrueTime={real_t}\n')
                        ASCfile.write(f'#AcqStart={date[0]}-{date[1]}-{date[2]}T{date[3]}:{date[4]}:{date[5]}\n')
                        DT=datetime.datetime(date[0],date[1],date[2],date[3],date[4],date[5])
                        RT=datetime.timedelta(0,real_t)
                        DT=str(DT+RT)
                        Da,Ti=str.split(DT)
                        ASCfile.write(f'#AcqEnd={Da}T{Ti}\n')
                        ASCfile.write('#Comment=\n')
                        ASCfile.write(f'#Title={sht}\n')
                        ASCfile.write('#FileName=\n')
                        ASCfile.write('#LinEnergyCalParams=\n')
                        ASCfile.write('#EnergyCalEquation=\n')
                        ASCfile.write('#FwhmCalParams=\n')
                        ASCfile.write('#FwhmCalEquation=\n')
                        ASCfile.write('#SpePartType[0]=\n')
                TLHPL.destroy()
                print(f'{len(ss)} spectra converted')
            except:
                print('unable to convert the provided file')
                
    def convert_CRMs_to_presets():
        for cert in NAA.certificates.keys():
            if not os.path.exists(f'data/presets/{cert}.spl'):
                with open(f'data/presets/{cert}.spl', 'w') as f:
                    for elem in NAA.certificates[cert].keys():
                        for item in A:
                            if item[0] == elem:
                                f.write(f'{item[1]} {item[2]}\n')
                
    def do_something(filename,noncertified_uncertainty):
        def linefeed(fs,i,default=0):
            vs = fs.row_values(i,0,5)
            try:
                float(vs[1])
            except:
                return (vs[0],None)
            else:
                if float(vs[1]) == 0:
                    return (vs[0],None)
                else:
                    if vs[2] == '':
                        vs[2] = vs[1]*default
                    return (vs[0],tuple(vs[1:]))
        wb = xlrd.open_workbook('data/sources/'+filename)
        certificates = {}
        column_line = []
        for sname in wb.sheet_names():
            fs = wb.sheet_by_name(sname)
            values = {}
            column_line += fs.col_values(0,1,fs.nrows)
            for i in range(1,fs.nrows-1):
                a,b = linefeed(fs,i,noncertified_uncertainty)
                if b is not None:
                    values[a] = b
            certificates[sname] = values
        column_line = sorted(set(column_line))
        return certificates,column_line
        
    def software_information():
        print(f'R-LENA\nversion {VERSION} ({VERSION_DATE})\nauthor m.diluzio@inrim.it\n') #update this string in case of modifications
    
    #Settings
    settings='data/kimp0-01r.txl'
    database,CRMs,tolerance_energy,rows,unc_stats,unc_stats_forall,unc_stats_standard,noncertified_uncertainty = initialization(settings)
    A=emission_database(database,['analytical'])
    #Initialization Analysis Class
    NAA=RNAAnalysis()
    noncertified_uncertainty = int(noncertified_uncertainty)/100
    NAA.certificates,column_line = do_something(CRMs,noncertified_uncertainty)
    convert_CRMs_to_presets() # in NAA.certificates c'è la risposta
    NAA.elem_dataframe = pd.DataFrame(columns = column_line)
    NAA.statistical_uncertainty_limit = int(unc_stats)
    NAA.statistical_uncertainty_limit_standard = int(unc_stats_standard)
    dummy_variable = [True,False]
    NAA.statistical_uncertainty_limit_forall = dummy_variable[int(unc_stats_forall)]
    NAA.info = {'version':VERSION, "version_CRM":CRMs, "version_emissions":database}
    #software information
    software_information()

    #GUI
    
    def mainscreen(M,NAA):
        def settings_modifications(M):
            def save_destroy(P1,P2,P3,P4,M,cvar,P5,P6,P7,P8):
                f=open(settings,'w')
                f.write(P1.get())
                f.write('\n')
                f.write(P2.get())
                f.write('\n')
                f.write(str(float(P3.get())))
                f.write('\n')
                f.write(str(int(P4.get())))
                f.write('\n')
                f.write(str(int(P5.get())))
                f.write('\n')
                f.write(str(int(P6.get())))
                f.write('\n')
                f.write(str(int(P7.get())))
                f.write('\n')
                f.write(str(int(P8.get())))
                f.close()
                if cvar.get():
                    lstsel = [filesel for filesel in os.listdir('data/presets/') if filesel[-4:]=='.spl']
                    for filename in lstsel:
                        os.remove('data/presets/'+filename)
                M.destroy()
            
            def change_parameter(V,name,somthing,Scale,Label):
                Label.configure(text=str(Scale.get()))
                
            def change_parameter2(V,name,somthing,Scale,Label):
                Label.configure(text=f'{Scale.get()}%')
            
            PRM1,PRM2,PRM3,PRM4,PRM5,PRM6,PRM7,PRM8=initialization(settings)
            TSS=Toplevel()
            TSS.title('Settings')
            TSS.resizable(False,False)
            TSS.focus()
            L=Label(TSS).pack()
            F=Frame(TSS)
            L=Label(F, text='', width=1).pack(side=LEFT)
            L=Label(F, text='Relative database', width=28, anchor=W).pack(side=LEFT)
            VVV=os.listdir('data/rel_data')
            values=[]
            for kls in VVV:
                if kls[-4:]=='.xls' or kls[-5:]=='.xlsx':
                    values.append(kls)
            databases_omboBCP=ttk.Combobox(F, values=values, state='readonly', width=35)
            databases_omboBCP.pack(side=LEFT)
            databases_omboBCP.set(PRM1)
            L=Label(F, text='', width=1).pack(side=LEFT)
            F.pack(anchor=W, pady=2)
            
            F=Frame(TSS)
            L=Label(F, text='', width=1).pack(side=LEFT)
            L=Label(F, text='Certificates', width=28, anchor=W).pack(side=LEFT)
            VVV=os.listdir('data/sources')
            valuesCRM=[]
            for kls in VVV:
                if kls[-4:]=='.xls' or kls[-5:]=='.xlsx':
                    valuesCRM.append(kls)
            databasesCRM_omboBCP=ttk.Combobox(F, values=valuesCRM, state='readonly', width=35)
            databasesCRM_omboBCP.pack(side=LEFT)
            databasesCRM_omboBCP.set(PRM2)
            L=Label(F, text='', width=1).pack(side=LEFT)
            F.pack(anchor=W, pady=2)
            
            F=Frame(TSS)
            L=Label(F, text='', width=1).pack(side=LEFT)
            L=Label(F, text='ΔE / keV', width=24, anchor=W).pack(side=LEFT)
            valueL=Label(F, text='', width=3, anchor=W)
            valueL.pack(side=LEFT)
            variable_energy_tolerance_S=DoubleVar(TSS)
            energy_tolerance_S = Scale(F, from_=0.1, to=2, orient=HORIZONTAL, resolution=0.1, width=15, length=150, variable=variable_energy_tolerance_S, showvalue=False)
            energy_tolerance_S.pack(side=LEFT)
            energy_tolerance_S.set(PRM3)
            valueL.configure(text=str(PRM3))
            event='w'
            variable_energy_tolerance_S.trace(event, lambda V='',name='',somthing='',Scale=variable_energy_tolerance_S,Label=valueL : change_parameter(V,name,somthing,Scale,Label))
            L=Label(F, text='', width=1).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(TSS)
            L=Label(F, text='', width=1).pack(side=LEFT)
            L=Label(F, text='N peak-list', width=24, anchor=W).pack(side=LEFT)
            valueL=Label(F, text='', width=3, anchor=W)
            valueL.pack(side=LEFT)
            variable_rows_S=IntVar(TSS)
            rows_S = Scale(F, from_=10, to=30, orient=HORIZONTAL, resolution=1, width=15, length=150, variable=variable_rows_S, showvalue=False)
            rows_S.pack(side=LEFT)
            rows_S.set(PRM4)
            valueL.configure(text=str(PRM4))
            variable_rows_S.trace(event, lambda V='',name='',somthing='',Scale=variable_rows_S,Label=valueL : change_parameter(V,name,somthing,Scale,Label))
            L=Label(F, text='', width=1).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(TSS)
            L=Label(F, text='', width=1).pack(side=LEFT)
            L=Label(F, text='Max peak uncertainty (sample)', width=24, anchor=W).pack(side=LEFT)
            valueLPS=Label(F, text='', width=3, anchor=W)
            valueLPS.pack(side=LEFT)
            variable_uncstats_S=IntVar(TSS)
            uncstats_S = Scale(F, from_=10, to=90, orient=HORIZONTAL, resolution=1, width=15, length=150, variable=variable_uncstats_S, showvalue=False)
            uncstats_S.pack(side=LEFT)
            uncstats_S.set(PRM5)
            valueLPS.configure(text=f'{PRM5}%')
            variable_uncstats_S.trace(event, lambda V='',name='',somthing='',Scale=variable_uncstats_S,Label=valueLPS : change_parameter2(V,name,somthing,Scale,Label))
            L=Label(F, text='', width=1).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(TSS)
            L=Label(F, text='', width=1).pack(side=LEFT)
            L=Label(F, text='Max peak uncertainty (std)', width=24, anchor=W).pack(side=LEFT)
            valueLPS_standard=Label(F, text='', width=3, anchor=W)
            valueLPS_standard.pack(side=LEFT)
            variable_uncstats_standard_S=IntVar(TSS)
            uncstats_standard_S = Scale(F, from_=5, to=15, orient=HORIZONTAL, resolution=1, width=15, length=150, variable=variable_uncstats_standard_S, showvalue=False)
            uncstats_standard_S.pack(side=LEFT)
            uncstats_standard_S.set(PRM7)
            valueLPS_standard.configure(text=f'{PRM7}%')
            variable_uncstats_standard_S.trace(event, lambda V='',name='',somthing='',Scale=variable_uncstats_standard_S,Label=valueLPS_standard : change_parameter2(V,name,somthing,Scale,Label))
            L=Label(F, text='', width=1).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(TSS)
            L=Label(F, text='', width=1).pack(side=LEFT)
            variable_extend_to_all = IntVar()
            L=Label(F, text='Set max peak uncertainty also for identified', width=40, anchor=W).pack(side=LEFT)
            RB01all = Radiobutton(F, text='True', variable=variable_extend_to_all, value=0)
            RB01all.pack(side=LEFT)
            RB02all = Radiobutton(F, text='False', variable=variable_extend_to_all, value=1)
            RB02all.pack(side=LEFT)
            variable_extend_to_all.set(int(PRM6))
            L=Label(F, text='', width=1).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(TSS)
            L=Label(F, text='', width=1).pack(side=LEFT)
            L=Label(F, text='Set non-certified uncertainty', width=24, anchor=W).pack(side=LEFT)
            valuenoncertified_standard=Label(F, text='', width=3, anchor=W)
            valuenoncertified_standard.pack(side=LEFT)
            variable_noncertified_standard_S=IntVar(TSS)
            noncertified_standard_S = Scale(F, from_=5, to=15, orient=HORIZONTAL, resolution=1, width=15, length=150, variable=variable_noncertified_standard_S, showvalue=False)
            noncertified_standard_S.pack(side=LEFT)
            noncertified_standard_S.set(PRM8)
            valuenoncertified_standard.configure(text=f'{PRM8}%')
            variable_noncertified_standard_S.trace(event, lambda V='',name='',somthing='',Scale=variable_noncertified_standard_S,Label=valuenoncertified_standard : change_parameter2(V,name,somthing,Scale,Label))
            L=Label(F, text='', width=1).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(TSS)
            L=Label(F, text='', width=1).pack(side=LEFT)
            CBUSL_bvar = BooleanVar()
            CBUSL = Checkbutton(F, variable=CBUSL_bvar, text='Update selection lists', width=28, anchor=W)
            CBUSL.pack(side=LEFT)
            CBUSL_bvar.set(False)
            F.pack(anchor=W, pady=5)
            B=Button(TSS, text='Confirm and close', command=lambda P1=databases_omboBCP,P2=databasesCRM_omboBCP,P3=variable_energy_tolerance_S,P4=variable_rows_S,M=M,cvar=CBUSL_bvar,P5=variable_uncstats_S,P6=variable_extend_to_all,P7=variable_uncstats_standard_S,P8=variable_noncertified_standard_S : save_destroy(P1,P2,P3,P4,M,cvar,P5,P6,P7,P8)).pack()
            L=Label(TSS).pack()
        
        def create_HyperLabmanually():
            def selectionLP(LP):
                types=[('Excel file','.xls'),('Excel file','.xlsx')]
                filename=askopenfilename(filetypes=types)
                LP.configure(text=filename)
                LP.focus()
            TLHPL=Toplevel()
            TLHPL.title('Select and convert')
            TLHPL.resizable(False,False)
            L=Label(TLHPL).pack()
            F=Frame(TLHPL)
            L=Label(F, width=1).pack(side=LEFT)
            BL=Button(F, text='Select', width=8)
            BL.pack(side=LEFT)
            L=Label(F, width=1).pack(side=LEFT)
            LPfile=Label(F, text='', width=60, anchor=E)
            LPfile.pack(side=LEFT)
            BL.configure(command=lambda LP=LPfile : selectionLP(LP))
            L=Label(F, width=1).pack(side=LEFT)
            F.pack(anchor=W)
            L=Label(TLHPL).pack()
            F=Frame(TLHPL)
            L=Label(F, width=1).pack(side=LEFT)
            BCv=Button(F, text='Convert', width=8, command=lambda wbname=LPfile,TLHPL=TLHPL : write_hyperlabfiles(wbname,TLHPL))
            BCv.pack(side=LEFT)
            F.pack(anchor=W)
            L=Label(TLHPL).pack()
            TLHPL.focus()
        
        def openpeaklistandspectrum(BT,NAA,LBB,LPKS,LDT):
            if BT.cget('text')=='Background':
                nomeHyperLabfile,startcounting,realT,liveT,peak_list,spectrum_counts,linE,linW=searchforalternateopenfile(NAA.statistical_uncertainty_limit,NAA.statistical_uncertainty_limit_forall,BT.cget('text'))
                if realT!=None:
                    S=Spectrum(BT.cget('text'),startcounting,realT,liveT,peak_list,spectrum_counts,nomeHyperLabfile)
                    NAA.set_backgroungspectrum(S)
                    shortnm=str.split(nomeHyperLabfile,'/')
                    LBB.configure(text=shortnm[-1])
                    LPKS.configure(text=str(len(S.peak_list)))
                    LDT.configure(text=S.readable_datetime())
            else:
                if BT.cget('text')=='Standard':
                    peak_limit_unc = NAA.statistical_uncertainty_limit_standard
                    NAAstatistical_uncertainty_limit_forall = True
                else:
                    peak_limit_unc = NAA.statistical_uncertainty_limit
                    NAAstatistical_uncertainty_limit_forall = NAA.statistical_uncertainty_limit_forall
                AAA=searchforalternateopenmultiplefiles(peak_limit_unc,NAAstatistical_uncertainty_limit_forall,BT.cget('text'))
                for item in AAA:
                    if item[2]!=None:
                        S=Spectrum(BT.cget('text'),item[1],item[2],item[3],item[4],item[5],item[0])
                        if S.define()=='Sample':
                            NAA.set_samplespectrum(S)
                            shortnm=str.split(item[0],'/')
                        else:
                            NAA.set_comparatorspectrum(S)
                            NAA.elem_dataframe.loc[item[0]] = np.nan
                            shortnm=str.split(item[0],'/')
                try:
                    if BT.cget('text')=='Sample':
                        if len(NAA.sample)>1:
                            LBB.configure(text=str(len(NAA.sample))+' spectra')
                            LPKS.configure(text='-')
                            LDT.configure(text='-')
                        else:
                            LBB.configure(text=shortnm[-1])
                            LPKS.configure(text=str(len(NAA.sample[0].peak_list)))
                            LDT.configure(text=NAA.sample[0].readable_datetime())
                    else:
                        if len(NAA.comparator)>1:
                            LBB.configure(text=str(len(NAA.comparator))+' spectra')
                            LPKS.configure(text='-')
                            LDT.configure(text='-')
                        else:
                            LBB.configure(text=shortnm[-1])
                            LPKS.configure(text=str(len(NAA.comparator[0].peak_list)))
                            LDT.configure(text=NAA.comparator[0].readable_datetime())
                except TypeError:
                    pass
                except UnboundLocalError:
                    pass
                
        def irradiation_info(BIR,NAA,LCH,LF,LALF,LIDT,LITM):
            def automatic_f_a(event,CB,F,UF,A,UA,CHNL):
                for t in CHNL:
                    if t[0]==CB.get():
                        try:
                            F.delete(0,END)
                            F.insert(0,t[1])
                            UF.delete(0,END)
                            UF.insert(0,t[2])
                            A.delete(0,END)
                            A.insert(0,t[3])
                            UA.delete(0,END)
                            UA.insert(0,t[4])
                            break
                        except:
                            pass
                        
            def automatic_irradiation_data(event,CB,DS,MS,YS,HS,MinS,SS,ITS,UTS,CHNL,CBB,F,UF,A,UA):
                for t in CHNL:
                    if t[0]==CB.get():
                        try:
                            DS.delete(0,END)
                            DS.insert(0,t[1])
                            MS.delete(0,END)
                            MS.insert(0,t[2])
                            YS.delete(0,END)
                            YS.insert(0,t[3])
                            HS.delete(0,END)
                            HS.insert(0,t[4])
                            MinS.delete(0,END)
                            MinS.insert(0,t[5])
                            SS.delete(0,END)
                            SS.insert(0,t[6])
                            ITS.delete(0,END)
                            ITS.insert(0,t[7])
                            UTS.delete(0,END)
                            UTS.insert(0,t[8])
                            CBB.set(t[9])
                            F.delete(0,END)
                            F.insert(0,t[10])
                            UF.delete(0,END)
                            UF.insert(0,t[11])
                            A.delete(0,END)
                            A.insert(0,t[12])
                            UA.delete(0,END)
                            UA.insert(0,t[13])
                            break
                        except:
                            pass
                        
            def save_update_channels(CB,F,UF,A,UA,CHNL):
                try:
                    float(F.get())
                    float(UF.get())
                    float(A.get())
                    float(UA.get())
                except:
                    pass
                else:
                    if float(F.get())>0 and CB.get()!='':
                        index=None
                        CBG=CB.get().replace(' ','_')
                        for th in range(len(CHNL)):
                            if CBG==CHNL[th][0]:
                                index=th
                                break
                        if index==None:
                            CHNL.append([CBG,float(F.get()),float(UF.get()),float(A.get()),float(UA.get())])
                        else:
                            CHNL[index]=[CBG,float(F.get()),float(UF.get()),float(A.get()),float(UA.get())]
                        listCH=[]
                        f=open('data/channels/chn.txt','w')
                        f.write('#flux parameters{this line is for comments and is not considered} #name_of_facility(without spaces) f uf alpha ualpha')
                        for th in CHNL:
                            listCH.append(th[0])
                            f.write('\n')
                            f.write(th[0])
                            f.write(' ')
                            f.write(str(th[1]))
                            f.write(' ')
                            f.write(str(th[2]))
                            f.write(' ')
                            f.write(str(th[3]))
                            f.write(' ')
                            f.write(str(th[4]))
                        f.close()
                        CB['values']=listCH
                        CB.set(CBG)
                        
            def delete_channel(CB,F,UF,A,UA,CHNL):
                if CB.get()!='':
                    CBG=CB.get().replace(' ','_')
                    for th in range(len(CHNL)):
                        if CBG==CHNL[th][0]:
                            CHNL.pop(th)
                            break
                    listCH=[]
                    f=open('data/channels/chn.txt','w')
                    f.write('#flux parameters{this line is for comments and is not considered} #name_of_facility(without spaces) f uf alpha ualpha')
                    for th in CHNL:
                        listCH.append(th[0])
                        f.write('\n')
                        f.write(th[0])
                        f.write(' ')
                        f.write(str(th[1]))
                        f.write(' ')
                        f.write(str(th[2]))
                        f.write(' ')
                        f.write(str(th[3]))
                        f.write(' ')
                        f.write(str(th[4]))
                    f.close()
                    CB['values']=listCH
                    CB.set('')
                    F.delete(0,END)
                    F.insert(END,'0.0')
                    UF.delete(0,END)
                    UF.insert(END,'0.0')
                    A.delete(0,END)
                    A.insert(END,'0.0000')
                    UA.delete(0,END)
                    UA.insert(END,'0.0000')
        
            def openchannels(n):
                """Retrieve f and α information from file"""
                f=open(n,'r')
                rl=f.readlines()
                f.close()
                rl.pop(0)
                r=[]
                for i in rl:
                    r.append(i.replace('\n',''))
                R=[]
                for i in r:
                    R.append(str.split(i))
                listR=[]
                for k in range(len(R)):
                    listR.append(R[k][0])
                    try:
                        R[k][1],R[k][2],R[k][3],R[k][4]=float(R[k][1]),float(R[k][2]),float(R[k][3]),float(R[k][4])
                    except:
                        pass
                return R,listR
            
            def openchannelsI(n):
                """Retrieve irradiation data information from file"""
                f=open(n,'r')
                rl=f.readlines()
                f.close()
                rl.pop(0)
                r=[]
                for i in rl:
                    r.append(i.replace('\n',''))
                R=[]
                for i in r:
                    R.append(str.split(i))
                listR=[]
                for k in range(len(R)):
                    listR.append(R[k][0])
                    try:
                        R[k][1],R[k][2],R[k][3],R[k][4],R[k][5],R[k][6],R[k][7],R[k][8],R[k][10],R[k][11],R[k][12],R[k][13]=float(R[k][1]),float(R[k][2]),float(R[k][3]),float(R[k][4]),float(R[k][5]),float(R[k][6]),float(R[k][7]),float(R[k][8],float(R[k][10]),float(R[k][11]),float(R[k][12]),float(R[k][13]))
                    except:
                        pass
                return R,listR
            
            def delete_irradiation_code(ECC,IRRD,listIRR):
                if ECC.get()!='':
                    index=None
                    try:
                        for i in range(len(IRRD)):
                            if ECC.get()==IRRD[i][0]:
                                index=i
                                break
                    except:
                        index=None
                    if index!=None:
                        IRRD.pop(index)
                        listIRR.pop(index)
                        ECC['values']=listIRR
                        ECC.set('')
                        f=open('data/irradiations/irr.txt','w')
                        f.write('#Irradiation data {This line is for comments and is ignored}')
                        for t in IRRD:
                            f.write('\n')
                            f.write(t[0])
                            f.write(' ')
                            f.write(str(t[1]))
                            f.write(' ')
                            f.write(str(t[2]))
                            f.write(' ')
                            f.write(str(t[3]))
                            f.write(' ')
                            f.write(str(t[4]))
                            f.write(' ')
                            f.write(str(t[5]))
                            f.write(' ')
                            f.write(str(t[6]))
                            f.write(' ')
                            f.write(str(t[7]))
                            f.write(' ')
                            f.write(str(t[8]))
                            f.write(' ')
                            f.write(t[9])
                            f.write(' ')
                            f.write(str(t[10]))
                            f.write(' ')
                            f.write(str(t[11]))
                            f.write(' ')
                            f.write(str(t[12]))
                            f.write(' ')
                            f.write(str(t[13]))
                        f.close()
            
            def apply(ECC,CHomboB,Fspinbox,UFspinbox,Aspinbox,UAspinbox,Dayspinbox,Monthspinbox,Yearspinbox,Hourspinbox,Minutespinbox,Secondspinbox,Itimespinbox,UItimespinbox,NAA,LCH,LF,LALF,LIDT,LITM,TL,IRRD):
                try:
                    float(Fspinbox.get())
                    float(UFspinbox.get())
                    float(Aspinbox.get())
                    float(UAspinbox.get())
                    int(Itimespinbox.get())
                    float(UItimespinbox.get())
                    dt=datetime.datetime(int(Yearspinbox.get()),int(Monthspinbox.get()),int(Dayspinbox.get()),int(Hourspinbox.get()),int(Minutespinbox.get()),int(Secondspinbox.get()))
                except:
                    print('Invalid data entered\na) f,a,u(f),u(a) and u(ti) should be of floating point type\nb) ti, should be of integer type\nc) set f and ti different from 0')
                else:
                    if float(Fspinbox.get())>0 and int(Itimespinbox.get())>0 and CHomboB.get()!='':
                        IR=Irradiation(dt,int(Itimespinbox.get()),float(UItimespinbox.get()),float(Fspinbox.get()),float(UFspinbox.get()),float(Aspinbox.get()),float(UAspinbox.get()),CHomboB.get(),ECC.get())
                        NAA.set_irradiation(IR)
                        LCH.configure(text=NAA.irradiation.channel)
                        LF.configure(text=str(NAA.irradiation.f))
                        LALF.configure(text=str(NAA.irradiation.a))
                        LIDT.configure(text=NAA.irradiation.readable_datetime())
                        LITM.configure(text=f'{NAA.irradiation.time}')
                        if ECC.get()!='':
                            ECCC=ECC.get().replace(' ','_')
                        else:
                            ECCC='_'
                        CHombo=CHomboB.get().replace(' ','_')
                        index=None
                        for th in range(len(IRRD)):
                            if ECCC==IRRD[th][0]:
                                index=th
                        if index==None:
                            IRRD.append([ECCC,int(Dayspinbox.get()),int(Monthspinbox.get()),int(Yearspinbox.get()),int(Hourspinbox.get()),int(Minutespinbox.get()),int(Secondspinbox.get()),int(Itimespinbox.get()),float(UItimespinbox.get()),CHombo,float(Fspinbox.get()),float(UFspinbox.get()),float(Aspinbox.get()),float(UAspinbox.get())])
                        else:
                            IRRD[index]=[ECCC,int(Dayspinbox.get()),int(Monthspinbox.get()),int(Yearspinbox.get()),int(Hourspinbox.get()),int(Minutespinbox.get()),int(Secondspinbox.get()),int(Itimespinbox.get()),float(UItimespinbox.get()),CHombo,float(Fspinbox.get()),float(UFspinbox.get()),float(Aspinbox.get()),float(UAspinbox.get())]
                        f=open('data/irradiations/irr.txt','w')
                        f.write('#Irradiation data {This line is for comments and is ignored}')
                        for t in IRRD:
                            f.write('\n')
                            f.write(t[0])
                            f.write(' ')
                            f.write(str(t[1]))
                            f.write(' ')
                            f.write(str(t[2]))
                            f.write(' ')
                            f.write(str(t[3]))
                            f.write(' ')
                            f.write(str(t[4]))
                            f.write(' ')
                            f.write(str(t[5]))
                            f.write(' ')
                            f.write(str(t[6]))
                            f.write(' ')
                            f.write(str(t[7]))
                            f.write(' ')
                            f.write(str(t[8]))
                            f.write(' ')
                            f.write(t[9])
                            f.write(' ')
                            f.write(str(t[10]))
                            f.write(' ')
                            f.write(str(t[11]))
                            f.write(' ')
                            f.write(str(t[12]))
                            f.write(' ')
                            f.write(str(t[13]))
                        f.close()
                        TL.destroy()
                    else:
                        print('Invalid data entered\na) f and ti should be greater than 0\nb) irr. channel name cannot be empty')
            
            CHNL,listCHN=openchannels('data/channels/chn.txt')
            IRRD,listIRR=openchannelsI('data/irradiations/irr.txt')
            if type(listCHN)!=list:
                listCHN=list(listCHN)
            TL=Toplevel()
            TL.title('Irradiation')
            TL.resizable(False,False)
            L=Label(TL).pack()
            F=Frame(TL)
            L=Label(F, width=1).pack(side=LEFT)
            L=Label(F, text='irr. code', width=9, anchor=W).pack(side=LEFT)
            CIRRComboB=ttk.Combobox(F, values=listIRR)
            CIRRComboB.pack(side=LEFT)
            try:
                CIRRComboB.set(NAA.imported_irradiation_code)
            except AttributeError:
                pass
            L=Label(F, width=1).pack(side=LEFT)
            BDELCIRR=Button(F, text='Delete', width=8, command=lambda CIRRComboB=CIRRComboB,IRRD=IRRD,listIRR=listIRR : delete_irradiation_code(CIRRComboB,IRRD,listIRR)).pack(side=LEFT)
            F.pack(anchor=W)
            L=Label(TL).pack()
            F=Frame(TL)
            L=Label(F, width=1).pack(side=LEFT)
            L=Label(F, text='irr. channel', width=9, anchor=W).pack(side=LEFT)
            CHomboB=ttk.Combobox(F, values=listCHN)
            CHomboB.pack(side=LEFT)
            L=Label(F, width=1).pack(side=LEFT)
            BSU=Button(F, text='Save', width=8)
            BSU.pack(side=LEFT)
            BDEL=Button(F, text='Delete', width=8)
            BDEL.pack(side=LEFT)
            L=Label(F, width=1).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(TL)
            L=Label(F, width=1).pack(side=LEFT)
            L=Label(F, text='f / 1', width=9, anchor=W).pack(side=LEFT)
            Fspinbox = Spinbox(F, from_=0, to=99999, width=10, increment=0.1)
            Fspinbox.pack(side=LEFT)
            L=Label(F, width=4).pack(side=LEFT)
            L=Label(F, text='u(f) / 1', width=6, anchor=W).pack(side=LEFT)
            UFspinbox = Spinbox(F, from_=0, to=99999, width=7, increment=0.1)
            UFspinbox.pack(side=LEFT)
            USH=Button(F, text='o', relief='flat', command=lambda L='Irradiation - u(f) / 1',u=UFspinbox : uncertainty_shaper(L,u))
            USH.pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(TL)
            L=Label(F, width=1).pack(side=LEFT)
            L=Label(F, text='α / 1', width=9, anchor=W).pack(side=LEFT)
            Aspinbox = Spinbox(F, from_=-1, to=1, width=10, increment=0.0001)
            Aspinbox.pack(side=LEFT)
            Aspinbox.delete(0,END)
            Aspinbox.insert(END,'0.0000')
            L=Label(F, width=4).pack(side=LEFT)
            L=Label(F, text='u(α) / 1', width=6, anchor=W).pack(side=LEFT)
            UAspinbox = Spinbox(F, from_=0.0, to=1, width=7, increment=0.0001)
            UAspinbox.pack(side=LEFT)
            UAspinbox.delete(0,END)
            UAspinbox.insert(END,'0.0000')
            USH=Button(F, text='o', relief='flat', command=lambda L='Irradiation - u(α) / 1',u=UAspinbox : uncertainty_shaper(L,u))
            USH.pack(side=LEFT)
            F.pack(anchor=W)
            BSU.configure(command=lambda CB=CHomboB,F=Fspinbox,UF=UFspinbox,A=Aspinbox,UA=UAspinbox,CHNL=CHNL:save_update_channels(CB,F,UF,A,UA,CHNL))
            BDEL.configure(command=lambda CB=CHomboB,F=Fspinbox,UF=UFspinbox,A=Aspinbox,UA=UAspinbox,CHNL=CHNL:delete_channel(CB,F,UF,A,UA,CHNL))
            L=Label(TL).pack()
            F=Frame(TL)
            L=Label(F, width=1).pack(side=LEFT)
            L=Label(F, text='ti / s', width=9, anchor=W).pack(side=LEFT)
            Itimespinbox = Spinbox(F, from_=0, to=999999999, width=10, increment=1)
            Itimespinbox.pack(side=LEFT)
            try:
                Itimespinbox.delete(0,END)
                Itimespinbox.insert(0,NAA.imported_lenght)
            except AttributeError:
                Itimespinbox.delete(0,END)
                Itimespinbox.insert(0,0)
            L=Label(F, width=4).pack(side=LEFT)
            L=Label(F, text='u(ti) / s', width=6, anchor=W).pack(side=LEFT)
            UItimespinbox = Spinbox(F, from_=0, to=1000, width=7, increment=1)
            UItimespinbox.pack(side=LEFT)
            USH=Button(F, text='o', relief='flat', command=lambda L='Irradiation - u(ti) / s',u=UItimespinbox : uncertainty_shaper(L,u))
            USH.pack(side=LEFT)
            F.pack(anchor=W)
            L=Label(TL).pack()
            F=Frame(TL)
            L=Label(F, width=1).pack(side=LEFT)
            L=Label(F, text='end irr.', width=6, anchor=W).pack(side=LEFT)
            L=Label(F, text='dd/MM/yyyy', width=11, anchor=W).pack(side=LEFT)
            Dayspinbox = Spinbox(F, from_=1, to=31, width=3, increment=1)
            Dayspinbox.pack(side=LEFT)
            try:
                Dayspinbox.delete(0,END)
                Dayspinbox.insert(0,NAA.imported_day)
            except AttributeError:
                Dayspinbox.delete(0,END)
                Dayspinbox.insert(0,1)
            Monthspinbox = Spinbox(F, from_=1, to=12, width=3, increment=1)
            Monthspinbox.pack(side=LEFT)
            try:
                Monthspinbox.delete(0,END)
                Monthspinbox.insert(0,NAA.imported_month)
            except AttributeError:
                Monthspinbox.delete(0,END)
                Monthspinbox.insert(0,1)
            Yearspinbox = Spinbox(F, from_=2019, to=2050, width=5, increment=1)
            Yearspinbox.pack(side=LEFT)
            try:
                Yearspinbox.delete(0,END)
                Yearspinbox.insert(0,NAA.imported_year)
            except AttributeError:
                Yearspinbox.delete(0,END)
                Yearspinbox.insert(0,2019)
            F.pack(anchor=W)
            F=Frame(TL)
            L=Label(F, width=8).pack(side=LEFT)
            L=Label(F, text='HH/mm/ss', width=11, anchor=W).pack(side=LEFT)
            Hourspinbox = Spinbox(F, from_=0, to=23, width=3, increment=1)
            Hourspinbox.pack(side=LEFT)
            try:
                Hourspinbox.delete(0,END)
                Hourspinbox.insert(0,NAA.imported_hours)
            except AttributeError:
                Hourspinbox.delete(0,END)
                Hourspinbox.insert(0,0)
            Minutespinbox = Spinbox(F, from_=0, to=59, width=3, increment=1)
            Minutespinbox.pack(side=LEFT)
            try:
                Minutespinbox.delete(0,END)
                Minutespinbox.insert(0,NAA.imported_minutes)
            except AttributeError:
                Minutespinbox.delete(0,END)
                Minutespinbox.insert(0,0)
            Secondspinbox = Spinbox(F, from_=0, to=59, width=3, increment=1)
            Secondspinbox.pack(side=LEFT)
            L=Label(F, width=1).pack(side=LEFT)
            F.pack(anchor=W)
            L=Label(TL).pack()
            event='<<ComboboxSelected>>'
            CHomboB.bind(event,lambda event=event,CB=CHomboB,F=Fspinbox,UF=UFspinbox,A=Aspinbox,UA=UAspinbox,CHNL=CHNL: automatic_f_a(event,CB,F,UF,A,UA,CHNL))
            try:
                CHomboB.set(NAA.imported_channelname)
                if NAA.imported_channelname in listCHN:
                    automatic_f_a(event,CHomboB,Fspinbox,UFspinbox,Aspinbox,UAspinbox,CHNL)
            except AttributeError:
                pass
            CIRRComboB.bind(event,lambda event=event,CB=CIRRComboB,DS=Dayspinbox,MS=Monthspinbox,YS=Yearspinbox,HS=Hourspinbox, MinS=Minutespinbox, SS=Secondspinbox, ITS=Itimespinbox, UTS=UItimespinbox, CHNL=IRRD,CBB=CHomboB,F=Fspinbox,UF=UFspinbox,A=Aspinbox,UA=UAspinbox: automatic_irradiation_data(event,CB,DS,MS,YS,HS,MinS,SS,ITS,UTS,CHNL,CBB,F,UF,A,UA))
            CB=Button(TL, text='Confirm', width=10, command= lambda ECC=CIRRComboB,CHomboB=CHomboB,Fspinbox=Fspinbox,UFspinbox=UFspinbox,Aspinbox=Aspinbox,UAspinbox=UAspinbox,Dayspinbox=Dayspinbox,Monthspinbox=Monthspinbox,Yearspinbox=Yearspinbox,Hourspinbox=Hourspinbox,Minutespinbox=Minutespinbox,Secondspinbox=Secondspinbox,Itimespinbox=Itimespinbox,UItimespinbox=UItimespinbox,NAA=NAA,LCH=LCH,LF=LF,LALF=LALF,LIDT=LIDT,LITM=LITM,TL=TL,IRRD=IRRD : apply(ECC,CHomboB,Fspinbox,UFspinbox,Aspinbox,UAspinbox,Dayspinbox,Monthspinbox,Yearspinbox,Hourspinbox,Minutespinbox,Secondspinbox,Itimespinbox,UItimespinbox,NAA,LCH,LF,LALF,LIDT,LITM,TL,IRRD))
            CB.pack()
            L=Label(TL).pack()
            TL.focus()
            
        def sparameters(mbox,EFC,ENE,FWH,vv=None):
            F=Frame(mbox)
            LNN=Frame(F)
            L=Label(LNN, text='', width=1, anchor=W).pack(side=LEFT)
            L=Label(LNN, text='ε param.', width=6, anchor=W).pack(side=LEFT)
            L=Label(LNN, text='x', width=8).pack(side=LEFT)
            L=Label(LNN, text='ur(x)', width=5).pack(side=LEFT)
            L=Label(LNN, width=10, anchor=W).pack(side=LEFT)
            if vv==None:
                L=Label(LNN, text='ε corr.', width=5).pack(side=LEFT)
                L=Label(LNN, text='a1', width=7).pack(side=LEFT)
                L=Label(LNN, text='a2', width=7).pack(side=LEFT)
                L=Label(LNN, text='a3', width=7).pack(side=LEFT)
                L=Label(LNN, text='a4', width=7).pack(side=LEFT)
                L=Label(LNN, text='a5', width=7).pack(side=LEFT)
                L=Label(LNN, text='a6', width=7).pack(side=LEFT)
                L=Label(LNN, text='', width=12).pack(side=LEFT)
                L=Label(LNN, text='E param.', width=6, anchor=W).pack(side=LEFT)
            LNN.pack(anchor=W)
            for xd in range(6):
                LNN=Frame(F)
                L=Label(LNN, text='', width=1, anchor=W).pack(side=LEFT)
                L=Label(LNN, text=f'a{xd+1}:', width=6, anchor=W).pack(side=LEFT)
                L=Label(LNN, text=f'{EFC.P[xd,0].__round__(6)}', width=8, anchor=W).pack(side=LEFT)
                if EFC.P[xd,0]!=0.0:
                    L=Label(LNN, text=f'({int(abs(EFC.P[xd,1]/EFC.P[xd,0])*100)}%)', width=5, anchor=W).pack(side=LEFT)
                else:
                    L=Label(LNN, text='(- %)', width=5, anchor=W).pack(side=LEFT)
                L=Label(LNN, width=10).pack(side=LEFT)
                if vv==None:
                    L=Label(LNN, text=f'a{xd+1}', width=5, anchor=W).pack(side=LEFT)
                    L=Label(LNN, text=f'{EFC.C[xd,0].__round__(4)}', width=7).pack(side=LEFT)
                    L=Label(LNN, text=f'{EFC.C[xd,1].__round__(4)}', width=7).pack(side=LEFT)
                    L=Label(LNN, text=f'{EFC.C[xd,2].__round__(4)}', width=7).pack(side=LEFT)
                    L=Label(LNN, text=f'{EFC.C[xd,3].__round__(4)}', width=7).pack(side=LEFT)
                    L=Label(LNN, text=f'{EFC.C[xd,4].__round__(4)}', width=7).pack(side=LEFT)
                    L=Label(LNN, text=f'{EFC.C[xd,5].__round__(4)}', width=7).pack(side=LEFT)
                    L=Label(LNN, text='', width=12).pack(side=LEFT)
                    if xd==0:
                        L=Label(LNN, text='b1:', width=3, anchor=W).pack(side=LEFT)
                        L=Label(LNN, text=f'{ENE.m.__round__(5)}', width=7, anchor=W).pack(side=LEFT)
                    if xd==1:
                        L=Label(LNN, text='b2:', width=3, anchor=W).pack(side=LEFT)
                        L=Label(LNN, text=f'{ENE.q.__round__(5)}', width=7, anchor=W).pack(side=LEFT)
                    if xd==3:
                        L=Label(LNN, text='FWHM param.', width=12, anchor=W).pack(side=LEFT)
                    if xd==4:
                        L=Label(LNN, text='c1:', width=3, anchor=W).pack(side=LEFT)
                        L=Label(LNN, text=f'{FWH.m.__round__(5)}', width=7, anchor=W).pack(side=LEFT)
                    if xd==5:
                        L=Label(LNN, text='c2:', width=3, anchor=W).pack(side=LEFT)
                        L=Label(LNN, text=f'{FWH.q.__round__(5)}', width=7, anchor=W).pack(side=LEFT)
                LNN.pack(anchor=W)
            L=Label(F).pack()
            F.pack(anchor=W)
            
        def showfit(box,lab):
            """Show fits"""
            if box.get()!='':
                x=np.linspace(58,3100,3043)
                ch_x=np.linspace(0,16000,4001)
                EEEE=lab.enegycomparatorfit.fun(ch_x)
                FFFF=lab.fwhmcomparatorfit.fun_squared(ch_x)
                GG=lab.efficiencycomparatorfit.fun(x)
                TL=Toplevel()
                TL.title('Calibration result - '+box.get())
                f = Figure(figsize=(8.5, 4))
                ax=f.add_subplot(221)
                Figur=Frame(TL)
                Figur.pack(anchor=CENTER, fill=BOTH, expand=1)
                canvas = FigureCanvasTkAgg(f, master=Figur)
                canvas.draw()
                canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=1)
                toolbar = NavigationToolbar2TkAgg(canvas, Figur)
                toolbar.update()
                ax.plot(x, GG, marker='', linestyle='-', color='b', linewidth=0.5)
                ax.set_ylabel(r'$\varepsilon$ / 1')
                ax.set_xlabel(r'$E$ / keV')
                ax.format_coord = lambda x, y: f'energy / keV={x.__round__(2)}, efficiency={y.__round__(6)}'
                ax.set_xlim(0,3000)
                if ax.get_ylim()[1]>1:
                    ax.set_ylim(0,0.5)
                ax.grid(linestyle='-.')
                axb=f.add_subplot(223)
                axb.plot(lab.dercomparatorfit[0], lab.dercomparatorfit[1], marker='', linestyle='-', color='b', linewidth=0.5)
                axb.grid(linestyle='-.')
                axb.set_ylabel(r'$\delta\varepsilon_r$ / mm$^{-1}$')
                axb.set_xlim(0,3000)
                axb.set_xlabel(r'$E$ / keV')
                axb.format_coord = lambda x, y: f'energy / keV={x.__round__(2)}, de={y.__round__(6)}'
                axey=f.add_subplot(222)
                axey.plot(ch_x,EEEE, marker='', linestyle='-', color='k', linewidth=0.5)
                axey.set_ylabel(r'$E$ / keV')
                axey.set_xlabel(r'channel / ch')
                axey.set_ylim(0,3000)
                xlimx=lab.enegycomparatorfit.fun_rev(3000)
                axey.set_xlim(0,xlimx)
                axey.grid(linestyle='-.')
                axey.format_coord = lambda x, y: f'channel={int(x)}, energy / keV={y.__round__(2)}'
                axfw=f.add_subplot(224)
                axfw.plot(ch_x,FFFF, marker='', linestyle='-', color='k', linewidth=0.5)
                axfw.set_ylabel(r'FWHM / ch')
                axfw.set_xlabel(r'channel / ch')
                axfw.set_xlim(0,xlimx)
                axfw.set_ylim(None,lab.fwhmcomparatorfit.fun_squared(xlimx))
                axfw.grid(linestyle='-.')
                axfw.format_coord = lambda x, y: f'channel={int(x)}, FWHM / ch={y.__round__(2)}'
                f.tight_layout()
                canvas.draw()
                sparameters(TL,lab.efficiencycomparatorfit,lab.enegycomparatorfit,lab.fwhmcomparatorfit)
                            
        def showspectrum(SC,S,B=None):
            def becamelog(ax,canvas,B):
                """Logarithm view of graph function"""
                if ax.get_yscale()=='linear':
                    ax.set_yscale('log', nonposy='clip')
                    B.configure(text='lin')
                else:
                    ax.set_yscale('linear')
                    B.configure(text='log')
                canvas.draw()
                
            def reset_command(ax,canvas,axis):
                """Reset the original view"""
                ax.axis(axis)
                canvas.draw()
                
            def on_scroll(event,ax,canvas,evx):
                """Scroll spectrum"""
                if event.xdata!=None and event.ydata!=None:
                    evx=evx.get()
                    if int(ax.get_xlim()[1]-ax.get_xlim()[0])!=evx:
                        try:
                            ax.set_xlim(int(event.xdata-evx/2),int(event.xdata+evx/2))
                        except ValueError:
                            pass
                        else:
                            canvas.draw()
                    elif int(ax.get_xlim()[1]-ax.get_xlim()[0])==evx:
                        try:
                            ax.set_xlim(int(ax.get_xlim()[0]+event.step*evx/6),int(ax.get_xlim()[1]+event.step*evx/6))
                        except ValueError:
                            pass
                        else:
                            canvas.draw()
                            
            def change_xticks(ax,canvas,B):
                """Change visualization unit of x axis"""
                if B.cget('text')=='keV':
                    ax.set_xticklabels(ticklabels)
                    ax.set_xlabel(r'$E$ / keV')
                    B.configure(text='ch')
                elif B.cget('text')=='ch':
                    ax.set_xticklabels(ticksint)
                    ax.set_xlabel('channel')
                    B.configure(text='keV')
                canvas.draw()
            
            def aggiungibackgroundinsicurezza(ax,canvas,S,B):
                """Add/drop background to the graph"""
                countsofthebackground=np.array(NAA.background.counts)
                countsofthebackground=(S.live_time/NAA.background.live_time)*countsofthebackground
                if len(ax.lines)==2:
                    ax.lines.pop()
                    B.configure(text='+BG')
                else:
                    ax.plot(np.linspace(0.5,len(countsofthebackground)+0.5,num=len(countsofthebackground)), countsofthebackground, marker='', linestyle='-', color='g', linewidth=0.5)
                    B.configure(text='-BG')
                canvas.draw()
                
            if S.counts is not None:#!=None:
                class CustomToolbar(NavigationToolbar2TkAgg):
                    toolitems = filter(lambda x: x[0] != "Home", NavigationToolbar2TkAgg.toolitems)
                    toolitems = filter(lambda x: x[0] != "Back", toolitems)
                    toolitems = filter(lambda x: x[0] != "Forward", toolitems)
                TL=Toplevel(SC)
                TL.title('Spectrum')
                f = Figure(figsize=(6, 4))
                ax=f.add_subplot(111)
                ax.format_coord = lambda x, y: f'channel={int(x)}, counts={int(y)}'
                Figur=Frame(TL)
                Figur.pack(anchor=CENTER, fill=BOTH, expand=1)
                canvas = FigureCanvasTkAgg(f, master=Figur)
                canvas.draw()
                canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=1)
                toolbar = CustomToolbar(canvas, Figur)
                toolbar.update()
                ax.plot(np.linspace(0.5,len(S.counts)+0.5,num=len(S.counts)), S.counts, marker='o', linestyle='-', color='k', linewidth=0.5, markersize=3, markerfacecolor='r')
                ax.set_ylabel('counts')
                ax.set_xlabel('channel')
                ticks=[]
                ticksint=[]
                ticklabels=[]
                for tyll in S.peak_list:
                    ticks.append(float(tyll[4]))
                    ticksint.append(int(float(tyll[4]).__round__(0)))
                    ticklabels.append(float(tyll[6]).__round__(1))
                ax.set_xticks(ticks)
                ax.set_xticklabels(ticksint)
                ax.grid(linestyle='-.')
                ax.set_xlim(0,len(S.counts))
                ax.set_ylim(1,max(S.counts)*3+2)
                ax.set_yscale('log', nonposy='clip')
                axis=ax.axis()
                f.tight_layout()
                canvas.draw()
                RESET=Button(TL, text='Reset', width=5, command=lambda ax=ax, canvas=canvas, axis=axis : reset_command(ax,canvas,axis))
                RESET.pack(side=LEFT)
                LOG=Button(TL, text='lin', width=3)
                LOG.configure(command=lambda ax=ax, canvas=canvas, B=LOG : becamelog(ax,canvas,B))
                LOG.pack(side=LEFT)
                if NAA.background!=None and S.identity!='Background':
                    BG=Button(TL, text='+BG', width=3)
                    BG.configure(command=lambda ax=ax, canvas=canvas, S=S, B=BG : aggiungibackgroundinsicurezza(ax,canvas,S,B))
                    BG.pack(side=LEFT)
                ECH=Button(TL, text='keV', width=3)
                ECH.configure(command=lambda ax=ax,canvas=canvas,B=ECH : change_xticks(ax,canvas,B))
                ECH.pack(side=LEFT)
                L=Label(TL, text='+', width=3, anchor=E).pack(side=LEFT)
                evx=Scale(TL, from_=40, to=320, resolution=40, showvalue=0, orient=HORIZONTAL)
                evx.pack(side=LEFT)
                evx.set(120)
                L=Label(TL, text='-').pack(side=LEFT)
                    
                cid=canvas.mpl_connect('scroll_event', lambda event='scroll_event',ax=ax,canvas=canvas,evx=evx : on_scroll(event,ax,canvas,evx))
                
            else:
                print('spectrum profile not available')
                
        def altreemissioni(I,Ai,S,E,T):
            """Find other emissions"""
            listaltremissioni=[]
            listaltriisotopi=[]
            for t in A:
                if I==t[2] and Ai==t[3] and S==t[4] and E!=t[5]:
                    try:
                        PR=str((float(t[12])).__round__(1))+' %'
                    except ValueError:
                        PR=str(t[12])+' %'
                    lineemiss=str(t[5])+' keV,  '+PR
                    listaltremissioni.append(lineemiss)
                elif T==t[1] and Ai!=t[3]:
                    if t[4]==1.0:
                        g=''
                    else:
                        g='m'
                    lineemiss=str(t[2])+'-'+str(int(t[3]))+g+',  '+str(t[5])+' keV'
                    listaltriisotopi.append(lineemiss)
                elif T==t[1] and Ai==t[3] and S!=t[4]:
                    if t[4]==1.0:
                        g=''
                    else:
                        g='m'
                    lineemiss=str(t[2])+'-'+str(int(t[3]))+g+',  '+str(t[5])+' keV'
                    listaltriisotopi.append(lineemiss)
            return listaltremissioni,listaltriisotopi #to check
                
        def commandinfoemission(WN):
            """Emission information screen"""
            if WN.get()!='':
                isot, energy = WN.get().split()
                for i in A:
                    if i[1] == isot and str(i[2]) == energy:
                        result = i[3]
                        break
                unit = 's'
                if result > 10000:
                    result = result / 3600
                    unit = 'h'
                    if result > 100:
                        result = result / 24
                        unit = 'd'
                        if result > 1000:
                            result = result / 365.24
                            unit = 'y'
                h, w, x, y = WN.winfo_height(), WN.winfo_width(), WN.winfo_rootx(), WN.winfo_rooty()
                IN = Toplevel()
                IN.geometry(f'{w}x{h}+{x+w}+{y}')
                IN.overrideredirect(True)
                IN.resizable(False,False)
                L=Label(IN, text=f't½ : {round(result,2)} {unit}', anchor=W)
                L.pack(side=LEFT, fill=X, expand=True)
                IN.focus()
                event="<FocusOut>"
                IN.bind(event,lambda event=event,M=IN : M.destroy())
                    
        def delete_preset_peakselection(preset):
            pass
                
        def save_preset_peakselection(preset,NA):
            if preset.get()!='':
                workinglist=NA.assign_nuclide[:]
                while '' in workinglist:
                    workinglist.remove('')
                if workinglist!=[]:
                    with open('data/presets/'+preset.get()+'.spl','w') as f:
                        for wd in workinglist:
                            f.write(wd)
                            f.write('\n')
                    h=os.listdir('data/presets')
                    values=[]
                    for t in h:
                        if t[-4:]=='.spl':
                            values.append(t[:-4])
                    values.append('')
                    preset['values']=values
                else:
                    print('the current emissions selection is empty')
            else:
                print('the preset name field is empty')
                
        def propagate_selection(NA,NAS,CBS):
            workinglist=NA.assign_nuclide[:]
            while '' in workinglist:
                workinglist.remove('')
            workingenergylist=[]
            for i in workinglist:
                mid=str.split(i)
                workingenergylist.append(float(mid[-1]))
            if len(workingenergylist)>0:
                for i in NAS:
                    if i!=NA:
                        i.assign_nuclide=['']*len(i.peak_list)
                        for tiy in range(len(i.peak_list)):
                            if float(i.peak_list[tiy][6])>workingenergylist[0]-3 and float(i.peak_list[tiy][6])<workingenergylist[-1]+3:
                                for pp in range(len(workingenergylist)):
                                    if float(i.peak_list[tiy][6])-float(tolerance_energy)<workingenergylist[pp] and float(i.peak_list[tiy][6])+float(tolerance_energy)>workingenergylist[pp]:
                                        i.assign_nuclide[tiy]=workinglist[pp]
                                        break
                    if CBS.get()!='':
                        i.selected_certificate = CBS.get()
            else:
                print('no emissions have been selected in the current peak list')
                
        def recall_emission_selection(F,paginazionepicchi,indice,NA,preset):
            NA.selected_certificate=preset.get()
            NA.assign_nuclide=['']*len(NA.peak_list)
            if preset.get()!='':
                try:
                    with open('data/presets/'+preset.get()+'.spl','r') as f:
                        r=f.readlines()
                    for t in range(len(r)):
                        r[t]=r[t].replace('\n','')
                except:
                    print(f'the selected {preset.get()} preset does not exist')
                else:
                    workingenergylist=[]
                    for i in r:
                        mid=str.split(i)
                        workingenergylist.append(float(mid[-1]))
                    sortedwlist = sorted(workingenergylist)
                    sortedwlist = [sortedwlist[0],sortedwlist[-1]]
                    for tiy in range(len(NA.peak_list)):
                        if float(NA.peak_list[tiy][6])>sortedwlist[0]-3 and float(NA.peak_list[tiy][6])<sortedwlist[-1]+3:
                            for pp in range(len(workingenergylist)):
                                if float(NA.peak_list[tiy][6])-float(tolerance_energy)<workingenergylist[pp] and float(NA.peak_list[tiy][6])+float(tolerance_energy)>workingenergylist[pp]:
                                    NA.assign_nuclide[tiy]=r[pp]
                                    break
                    if NA.identity=='Standard':
                        sciogliipicchicom(F,paginazionepicchi,indice,NA)
                    else:
                        sciogliipicchiana(F,paginazionepicchi,indice,NA)
            else:
                if NA.identity=='Standard':
                    sciogliipicchicom(F,paginazionepicchi,indice,NA)
                else:
                    sciogliipicchiana(F,paginazionepicchi,indice,NA)
                    
        def funtion_autoselect_emission(NA,alternate_database=None):
            if alternate_database is not None:
                workingenergylist=[]
                for i in alternate_database:
                    workingenergylist.append(float(i[2]))
                sortedwlist = sorted(workingenergylist)
                sortedwlist = [sortedwlist[0],sortedwlist[-1]]
                for tiy in range(len(NA.peak_list)):
                    if float(NA.peak_list[tiy][6])>sortedwlist[0]-3 and float(NA.peak_list[tiy][6])<sortedwlist[-1]+3:
                        for pp in range(len(workingenergylist)):
                            if float(NA.peak_list[tiy][6])-float(tolerance_energy)<workingenergylist[pp] and float(NA.peak_list[tiy][6])+float(tolerance_energy)>workingenergylist[pp]:
                                NA.assign_nuclide[tiy]=f'{alternate_database[pp][1]} {str(alternate_database[pp][2])}'
                                break
            else:
                NA.assign_nuclide=['']*len(NA.peak_list)
                
        def singlescreen_of_multiple(NA,SC,NAS,alternate_database=None):
            """Display single of multiple spectra"""
            ratiooflines=int(rows)
            residuallenS=len(NA.peak_list)
            i=0
            paginazionepicchi=[]
            if NA.assign_nuclide==None:
                NA.assign_nuclide=['']*len(NA.peak_list)
            if NA.identity=='Sample':
                funtion_autoselect_emission(NA,alternate_database)
            while residuallenS>0:
                try:
                    paginapicco=NA.peak_list[ratiooflines*i:ratiooflines*i+ratiooflines]
                except IndexError:
                    paginapicco=NA.peak_list[ratiooflines*i:]
                paginazionepicchi.append(paginapicco)
                residuallenS=residuallenS-ratiooflines
                i=i+1
            cdn=SC.winfo_children()
            for i in cdn:
                i.destroy()
            F=Frame(SC)
            L=Label(F, width=1).pack(side=LEFT)
            BSS=Button(F, text='Plot', width=8)
            BSS.pack(side=LEFT, anchor=W)
            BSS.configure(command=lambda SC=SC,S=NA: showspectrum(SC,S))
            if NA.identity == 'Standard':
                L=Label(F, width=5).pack(side=LEFT)
                L=Label(F, text='selection name', width=12).pack(side=LEFT)
                lLD=os.listdir('data/presets')
                CLD=[]
                for gk in lLD:
                    if gk[-4:]=='.spl':
                        CLD.append(gk[:-4])
                CLD.append('')
                SP_comboB=ttk.Combobox(F, values=CLD, width=20, state='readonly')
                SP_comboB.pack(side=LEFT)
                try:
                    SP_comboB.set(NA.selected_certificate)
                except AttributeError:
                    pass
                BTTSSRC=Button(F, text='Recall', width=8)
                BTTSSRC.pack(side=LEFT, anchor=W)
                BTTSS=Button(F, text='Apply', width=8)
                BTTSS.pack(side=LEFT, anchor=W)
                BTTSS.configure(command=lambda NA=NA,NAS=NAS,CBS=SP_comboB: propagate_selection(NA,NAS,CBS))
            L=Label(F, width=1).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(SC)
            L=Label(F, width=1).pack(side=LEFT)
            L=Label(F, text='start acqusition: '+NA.readable_datetime(), width=30, anchor=W).pack(side=LEFT)
            L=Label(F, text='tc / s: '+str(NA.real_time), width=16, anchor=W).pack(side=LEFT)
            L=Label(F, text='tl / s: '+str(NA.live_time), width=16, anchor=W).pack(side=LEFT)
            L=Label(F, text='tdead: '+NA.deadtime(), width=14, anchor=W).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(SC)
            L=Label(F, width=1).pack(side=LEFT)
            if len(NA.spectrumpath)>71:
                L=Label(F, text='...'+NA.spectrumpath[-70:], width=76, anchor=E).pack(side=LEFT)
            else:
                L=Label(F, text=NA.spectrumpath, width=76, anchor=E).pack(side=LEFT)
            F.pack(anchor=W)
            L=Label(SC).pack()
            F=Frame(SC)
            P=Frame(F)
            P.pack()
            indice=IntVar(SC)
            indice.set(0)
            if NA.identity == 'Standard':
                BTTSSRC.configure(command=lambda F=F,paginazionepicchi=paginazionepicchi,indice=indice,NA=NA,preset=SP_comboB: recall_emission_selection(F,paginazionepicchi,indice,NA,preset))
                sciogliipicchicom(F,paginazionepicchi,indice,NA)
            else:
                sciogliipicchiana(F,paginazionepicchi,indice,NA,alternate_database)
            F.pack()
            
        def Bmenocommandscroll(MS,superi,NA,alternate_database=None):
            if superi>0:
                superi-=1
                cdn=MS.winfo_children()
                for i in cdn:
                    i.destroy()
                F=Frame(MS)
                Bmeno=Button(F, text='<', relief='flat', command=lambda MS=MS,superi=superi,NA=NA,alternate_database=alternate_database : Bmenocommandscroll(MS,superi,NA,alternate_database)).pack(side=LEFT)
                L=Label(F, text=f'spectrum {superi+1} of {len(NA)}').pack(side=LEFT)
                Bpiu=Button(F, text='>', relief='flat', command=lambda MS=MS,superi=superi,NA=NA,alternate_database=alternate_database : Bpiucommandscroll(MS,superi,NA,alternate_database)).pack(side=LEFT)
                F.pack()
                L=Label(MS).pack()
                SC=Frame(MS)
                SC.pack()
                singlescreen_of_multiple(NA[superi],SC,NAS=NA,alternate_database=alternate_database)
        
        def Bpiucommandscroll(MS,superi,NA,alternate_database=None):
            if superi<len(NA)-1:
                superi+=1
                cdn=MS.winfo_children()
                for i in cdn:
                    i.destroy()
                F=Frame(MS)
                Bmeno=Button(F, text='<', relief='flat', command=lambda MS=MS,superi=superi,NA=NA,alternate_database=alternate_database : Bmenocommandscroll(MS,superi,NA,alternate_database)).pack(side=LEFT)
                L=Label(F, text=f'spectrum {superi+1} of {len(NA)}').pack(side=LEFT)
                Bpiu=Button(F, text='>', relief='flat', command=lambda MS=MS,superi=superi,NA=NA,alternate_database=alternate_database : Bpiucommandscroll(MS,superi,NA,alternate_database)).pack(side=LEFT)
                F.pack()
                L=Label(MS).pack()
                SC=Frame(MS)
                SC.pack()
                singlescreen_of_multiple(NA[superi],SC,NAS=NA,alternate_database=alternate_database)
            
        def multiplescreen(NA,alternate_database=None):
            """Display multiple spectra"""
            superi=0
            MS=Toplevel()
            MS.title(NA[0].identity+' peak list')
            MS.resizable(False,False)
            MS.focus()
            F=Frame(MS)
            Bmeno=Button(F, text='<', relief='flat', command=lambda MS=MS,superi=superi,NA=NA,alternate_database=alternate_database : Bmenocommandscroll(MS,superi,NA,alternate_database)).pack(side=LEFT)
            L=Label(F, text=f'spectrum {superi+1} of {len(NA)}').pack(side=LEFT)
            Bpiu=Button(F, text='>', relief='flat', command=lambda MS=MS,superi=superi,NA=NA,alternate_database=alternate_database : Bpiucommandscroll(MS,superi,NA,alternate_database)).pack(side=LEFT)
            F.pack()
            L=Label(MS).pack()
            SC=Frame(MS)
            SC.pack()
            singlescreen_of_multiple(NA[superi],SC,NAS=NA,alternate_database=alternate_database)
        
        def search(energia,tz=0.3,alternate_database=None):
            """Search correspondance in k0 database"""
            result=[]
            tz=float(tz)
            energia=float(energia)
            if alternate_database is not None:
                for i in alternate_database:
                    try:
                        if i[2]+tz>energia and i[2]-tz<energia:
                            lines=i[1]+' '+str(i[2])
                            result.append(lines)
                    except:
                        pass
            else:
                for i in A:
                    try:
                        if i[2]+tz>energia and i[2]-tz<energia:
                            lines=i[1]+' '+str(i[2])
                            result.append(lines)
                    except:
                        pass
            return result
        
        def BmenocommandANA(P,paginazionepicchi,indice,NA,alternate_database=None):
            if indice.get()>0:
                indice.set(indice.get()-1)
                sciogliipicchiana(P,paginazionepicchi,indice,NA,alternate_database)
        
        def BpiucommandANA(P,paginazionepicchi,indice,NA,alternate_database=None):
            if indice.get()<len(paginazionepicchi)-1:
                indice.set(indice.get()+1)
                sciogliipicchiana(P,paginazionepicchi,indice,NA,alternate_database)
                
        def selectionecomboselectedWN(event,WN,indice,valuei,NAN):
            NAN[indice.get()*int(rows)+valuei]=WN.get()
            
        def eliminapasticci(event,WN,indice,paginazionepicchi,valuei,NAN):
            WN.set('')
            NAN[indice.get()*int(rows)+valuei]=WN.get()
        
        def sciogliipicchiana(P,paginazionepicchi,indice,NA,alternate_database=None):
            """Display peaklist analyte"""
            if indice.get()>-1 and indice.get()<len(paginazionepicchi):
                cdn=P.winfo_children()
                for i in cdn:
                    i.destroy()
                F=Frame(P)
                chrw=10
                Bmeno=Button(F, text='<', relief='flat', command=lambda P=P,paginazionepicchi=paginazionepicchi,indice=indice,NA=NA,alternate_database=alternate_database : BmenocommandANA(P,paginazionepicchi,indice,NA,alternate_database)).pack(side=LEFT)
                L=Label(F, text=f'page {indice.get()+1} of {len(paginazionepicchi)}').pack(side=LEFT)
                Bpiu=Button(F, text='>', relief='flat', command=lambda P=P,paginazionepicchi=paginazionepicchi,indice=indice,NA=NA,alternate_database=alternate_database : BpiucommandANA(P,paginazionepicchi,indice,NA,alternate_database)).pack(side=LEFT)
                F.pack()
                titles=['channel','E / keV','net area / 1','uncertainty','FWHM / ch','emission','selection','','','','','','']
                FL=Frame(P)
                L=Label(FL, text=titles[0], width=chrw).pack(side=LEFT)
                L=Label(FL, text=titles[1], width=chrw).pack(side=LEFT)
                L=Label(FL, text=titles[2], width=chrw).pack(side=LEFT) #area
                L=Label(FL, text=titles[3], width=chrw).pack(side=LEFT) #uncertainty
                L=Label(FL, text=titles[4], width=chrw).pack(side=LEFT) #FWHM
                L=Label(FL, text=titles[5], width=chrw+2).pack(side=LEFT)
                FL.pack(anchor=W)
                valuei=0
                for j in range(len(paginazionepicchi[indice.get()])):
                    FL=Frame(P)
                    L=Label(FL, text=str(float(paginazionepicchi[indice.get()][j][4]).__round__(2)), width=chrw).pack(side=LEFT) #channel
                    L=Label(FL, text=str(float(paginazionepicchi[indice.get()][j][6]).__round__(2)), width=chrw).pack(side=LEFT) #energy
                    L=Label(FL, text=str(int(float(paginazionepicchi[indice.get()][j][8]))), width=chrw).pack(side=LEFT) #area
                    dinc=float(paginazionepicchi[indice.get()][j][9])/float(paginazionepicchi[indice.get()][j][8])*100
                    L=Label(FL, text=str(dinc.__round__(2))+' %', width=chrw).pack(side=LEFT) #uncetainty
                    if float(paginazionepicchi[indice.get()][j][10])<0.01:
                        L=Label(FL, text='-', width=chrw).pack(side=LEFT) #FWHM
                    else:
                        L=Label(FL, text=paginazionepicchi[indice.get()][j][10], width=chrw).pack(side=LEFT) #FWHM
                    possiblenuclides=search(paginazionepicchi[indice.get()][j][6],tolerance_energy,alternate_database)
                    WN=ttk.Combobox(FL, values=possiblenuclides, state='readonly')
                    WN.set(NA.assign_nuclide[indice.get()*int(rows)+j])
                    WN.pack(side=LEFT)
                    WN.configure(width=chrw+2)
                    ewent='<<ComboboxSelected>>'
                    WN.bind(ewent, lambda event=ewent,WN=WN,indice=indice,valuei=valuei,NAN=NA.assign_nuclide: selectionecomboselectedWN(event,WN,indice,valuei,NAN))
                    if len(possiblenuclides)==0:
                        WN.configure(state='disabled')
                    if len(possiblenuclides)>0:
                        NboX=Label(FL, text=f'({len(possiblenuclides)})', width=3).pack(side=LEFT)
                    else:
                        NboX=Label(FL, text=f'', width=3).pack(side=LEFT)
                    XboX=Button(FL, text='X', width=3, command= lambda event=ewent,WN=WN,indice=indice,paginazionepicchi=paginazionepicchi,valuei=valuei,NAN=NA.assign_nuclide : eliminapasticci(event,WN,indice,paginazionepicchi,valuei,NAN))
                    XboX.pack(side=LEFT)
                    Buttoninfo=Button(FL, text='Info', width=chrw, command= lambda WN=WN: commandinfoemission(WN))
                    Buttoninfo.pack(side=LEFT)
                    L=Label(FL, width=1).pack(side=LEFT)
                    FL.pack(anchor=W)
                    valuei=valuei+1
                L=Label(P).pack()

        def Bmenocommandcom(P,paginazionepicchi,indice,NA):
            if indice.get()>0:
                indice.set(indice.get()-1)
                sciogliipicchicom(P,paginazionepicchi,indice,NA)
        
        def Bpiucommandcom(P,paginazionepicchi,indice,NA):
            if indice.get()<len(paginazionepicchi)-1:
                indice.set(indice.get()+1)
                sciogliipicchicom(P,paginazionepicchi,indice,NA)
                
        def sciogliipicchicom(P,paginazionepicchi,indice,NA):
            """Display peaklist comparator"""
            if indice.get()>-1 and indice.get()<len(paginazionepicchi):
                cdn=P.winfo_children()
                for i in cdn:
                    i.destroy()
                F=Frame(P)
                chrw=10
                Bmeno=Button(F, text='<', relief='flat', command=lambda P=P,paginazionepicchi=paginazionepicchi,indice=indice,NA=NA : Bmenocommandcom(P,paginazionepicchi,indice,NA)).pack(side=LEFT)
                L=Label(F, text=f'page {indice.get()+1} of {len(paginazionepicchi)}').pack(side=LEFT)
                Bpiu=Button(F, text='>', relief='flat', command=lambda P=P,paginazionepicchi=paginazionepicchi,indice=indice,NA=NA : Bpiucommandcom(P,paginazionepicchi,indice,NA)).pack(side=LEFT)
                F.pack()
                titles=['channel','E / keV','net area / 1','uncertainty','FWHM / ch','emission','selection','relative','','','','','']
                FL=Frame(P)
                #L=Label(FL, text=titles[6], width=chrw).pack(side=LEFT)
                L=Label(FL, text=titles[0], width=chrw).pack(side=LEFT)
                L=Label(FL, text=titles[1], width=chrw).pack(side=LEFT)
                L=Label(FL, text=titles[2], width=chrw).pack(side=LEFT) #area
                L=Label(FL, text=titles[3], width=chrw).pack(side=LEFT) #uncertainty
                L=Label(FL, text=titles[4], width=chrw).pack(side=LEFT) #FWHM
                L=Label(FL, text=titles[5], width=chrw+2).pack(side=LEFT)
                L=Label(FL, text=titles[8], width=chrw).pack(side=LEFT)
                FL.pack(anchor=W)
                valuei=0
                #CheckVar=StringVar(P)
                for j in range(len(paginazionepicchi[indice.get()])):
                    FL=Frame(P)
                    L=Label(FL, text=str(float(paginazionepicchi[indice.get()][j][4]).__round__(2)), width=chrw).pack(side=LEFT) # channel
                    L=Label(FL, text=str(float(paginazionepicchi[indice.get()][j][6]).__round__(2)), width=chrw).pack(side=LEFT) #energy
                    L=Label(FL, text=str(int(float(paginazionepicchi[indice.get()][j][8]))), width=chrw).pack(side=LEFT) #area
                    dinc=float(paginazionepicchi[indice.get()][j][9])/float(paginazionepicchi[indice.get()][j][8])*100
                    L=Label(FL, text=str(dinc.__round__(2))+' %', width=chrw).pack(side=LEFT) #uncertaonty
                    if float(paginazionepicchi[indice.get()][j][10])<0.01:
                        L=Label(FL, text='-', width=chrw).pack(side=LEFT) #FWHM
                    else:
                        L=Label(FL, text=paginazionepicchi[indice.get()][j][10], width=chrw).pack(side=LEFT) #FWHM
                    possiblenuclides=search(paginazionepicchi[indice.get()][j][6],tolerance_energy)
                    WN=ttk.Combobox(FL, values=possiblenuclides, state='readonly')
                    WN.set(NA.assign_nuclide[indice.get()*int(rows)+j])
                    WN.pack(side=LEFT)
                    WN.configure(width=chrw+2)
                    ewent='<<ComboboxSelected>>'
                    WN.bind(ewent, lambda event=ewent,WN=WN,indice=indice,valuei=valuei,NAN=NA.assign_nuclide: selectionecomboselectedWN(event,WN,indice,valuei,NAN))
                    if len(possiblenuclides)==0:
                        WN.configure(state='disabled')
                    if len(possiblenuclides)>0:
                        NboX=Label(FL, text=f'({len(possiblenuclides)})', width=3).pack(side=LEFT)
                    else:
                        NboX=Label(FL, text=f'', width=3).pack(side=LEFT)
                    XboX=Button(FL, text='X', width=3, command= lambda event=ewent,WN=WN,indice=indice,paginazionepicchi=paginazionepicchi,valuei=valuei,NAN=NA.assign_nuclide : eliminapasticci(event,WN,indice,paginazionepicchi,valuei,NAN))
                    XboX.pack(side=LEFT)
                    Buttoninfo=Button(FL, text='Info', width=chrw, command= lambda WN=WN: commandinfoemission(WN))
                    Buttoninfo.pack(side=LEFT)
                    L=Label(FL, width=1).pack(side=LEFT)
                    FL.pack(anchor=W)
                    valuei=valuei+1
                L=Label(P).pack()
                
        def Bmenocommandbkg(P,paginazionepicchi,indice):
            if indice.get()>0:
                indice.set(indice.get()-1)
                sciogliipicchibkg(P,paginazionepicchi,indice)
        
        def Bpiucommandbkg(P,paginazionepicchi,indice):
            if indice.get()<len(paginazionepicchi)-1:
                indice.set(indice.get()+1)
                sciogliipicchibkg(P,paginazionepicchi,indice)
        
        def sciogliipicchibkg(P,paginazionepicchi,indice):
            """Display peaklist background"""
            if indice.get()>-1 and indice.get()<len(paginazionepicchi):
                cdn=P.winfo_children()
                for i in cdn:
                    i.destroy()
                F=Frame(P)
                chrw=10
                Bmeno=Button(F, text='<', relief='flat', command=lambda P=P,paginazionepicchi=paginazionepicchi,indice=indice : Bmenocommandbkg(P,paginazionepicchi,indice)).pack(side=LEFT)
                L=Label(F, text=f'page {indice.get()+1} of {len(paginazionepicchi)}').pack(side=LEFT)
                Bpiu=Button(F, text='>', relief='flat', command=lambda P=P,paginazionepicchi=paginazionepicchi,indice=indice : Bpiucommandbkg(P,paginazionepicchi,indice)).pack(side=LEFT)
                F.pack()
                titles=['channel','E / keV','net area / 1','uncertainty','FWHM / ch','emission','selection','','','','','','']
                FL=Frame(P)
                L=Label(FL, text=titles[0], width=chrw).pack(side=LEFT)
                L=Label(FL, text=titles[1], width=chrw).pack(side=LEFT)
                L=Label(FL, text=titles[2], width=chrw).pack(side=LEFT) #area
                L=Label(FL, text=titles[3], width=chrw).pack(side=LEFT) #uncertainty
                L=Label(FL, text=titles[4], width=chrw).pack(side=LEFT) #FWHM
                FL.pack(anchor=W)
                for j in range(len(paginazionepicchi[indice.get()])):
                    FL=Frame(P)
                    L=Label(FL, text=str(float(paginazionepicchi[indice.get()][j][4]).__round__(2)), width=chrw).pack(side=LEFT) # channel
                    L=Label(FL, text=str(float(paginazionepicchi[indice.get()][j][6]).__round__(2)), width=chrw).pack(side=LEFT) #energy
                    L=Label(FL, text=str(int(float(paginazionepicchi[indice.get()][j][8]))), width=chrw).pack(side=LEFT) #area
                    dinc=float(paginazionepicchi[indice.get()][j][9])/float(paginazionepicchi[indice.get()][j][8])*100
                    L=Label(FL, text=str(dinc.__round__(2))+' %', width=chrw).pack(side=LEFT) #uncertainty
                    if float(paginazionepicchi[indice.get()][j][10])<0.01:
                        L=Label(FL, text='-', width=chrw).pack(side=LEFT) #FWHM
                    else:
                        L=Label(FL, text=paginazionepicchi[indice.get()][j][10], width=chrw).pack(side=LEFT) #FWHM
                    FL.pack(anchor=W)
                L=Label(P).pack()
            
        def singlescreen(NA,alternate_database=None):
            """Display single spectrum"""
            ratiooflines=int(rows)
            residuallenS=len(NA.peak_list)
            i=0
            paginazionepicchi=[]
            if NA.assign_nuclide==None:
                NA.assign_nuclide=['']*len(NA.peak_list)
            if NA.identity=='Sample':
                funtion_autoselect_emission(NA,alternate_database)
            while residuallenS>0:
                try:
                    paginapicco=NA.peak_list[ratiooflines*i:ratiooflines*i+ratiooflines]
                except IndexError:
                    paginapicco=NA.peak_list[ratiooflines*i:]
                paginazionepicchi.append(paginapicco)
                residuallenS=residuallenS-ratiooflines
                i=i+1
            SC=Toplevel()
            SC.title(NA.identity+' peak list')
            SC.resizable(False,False)
            SC.focus()
            F=Frame(SC)
            L=Label(SC).pack()
            L=Label(F, width=1).pack(side=LEFT)
            BSS=Button(F, text='Plot', width=8)
            BSS.pack(side=LEFT, anchor=W)
            BSS.configure(command=lambda SC=SC,S=NA: showspectrum(SC,S))
            if NA.identity == 'Standard':
                L=Label(F, width=5).pack(side=LEFT)
                L=Label(F, text='selection name', width=12).pack(side=LEFT)
                lLD=os.listdir('data/presets')
                CLD=[]
                for gk in lLD:
                    if gk[-4:]=='.spl':
                        CLD.append(gk[:-4])
                CLD.append('')
                SP_comboB=ttk.Combobox(F, values=CLD, width=20, state='readonly')
                SP_comboB.pack(side=LEFT)
                try:
                    SP_comboB.set(NA.selected_certificate)
                except AttributeError:
                    pass
                BTTSSRC=Button(F, text='Recall', width=8)
                BTTSSRC.pack(side=LEFT, anchor=W)
                L=Label(F, width=1).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(SC)
            L=Label(F, width=1).pack(side=LEFT)
            L=Label(F, text='start acqusition: '+NA.readable_datetime(), width=30, anchor=W).pack(side=LEFT)
            L=Label(F, text='tc / s: '+str(NA.real_time), width=16, anchor=W).pack(side=LEFT)
            L=Label(F, text='tl / s: '+str(NA.live_time), width=16, anchor=W).pack(side=LEFT)
            L=Label(F, text='tdead: '+NA.deadtime(), width=14, anchor=W).pack(side=LEFT)
            F.pack(anchor=W)
            F=Frame(SC)
            L=Label(F, width=1).pack(side=LEFT)
            if len(NA.spectrumpath)>71:
                L=Label(F, text='...'+NA.spectrumpath[-70:], width=76, anchor=E).pack(side=LEFT)
            else:
                L=Label(F, text=NA.spectrumpath, width=76, anchor=E).pack(side=LEFT)
            F.pack(anchor=W)
            L=Label(SC).pack()
            F=Frame(SC)
            P=Frame(F)
            P.pack()
            indice=IntVar(SC)
            indice.set(0)
            if NA.identity=='Standard':#here is the point
                sciogliipicchicom(F,paginazionepicchi,indice,NA)
                BTTSSRC.configure(command=lambda F=F,paginazionepicchi=paginazionepicchi,indice=indice,NA=NA,preset=SP_comboB: recall_emission_selection(F,paginazionepicchi,indice,NA,preset))
            elif NA.identity=='Sample':
                sciogliipicchiana(F,paginazionepicchi,indice,NA,alternate_database)
            else:
                sciogliipicchibkg(F,paginazionepicchi,indice)
            F.pack()
            
        def get_master_emissions(NAA):
            if NAA.comparator is not None:
                p_list = []
                for spectrum in NAA.comparator:
                    if spectrum.assign_nuclide is not None:
                        add = [item for item in spectrum.assign_nuclide if item!='' and item not in p_list]
                        p_list += add
                if len(p_list)==0:
                    return None
                else:
                    A_shortlist = []
                    for line in p_list:
                        emit,energ = line.split()
                        for A_line in A:
                            if emit == A_line[1] and energ == str(A_line[2]):
                                A_shortlist.append(A_line)
                                break
                    return A_shortlist
            return None
            
        def overlookscreen(NA,alternate_database=None):
            if NA!=None:
                try:
                    len(NA)
                except TypeError:
                    singlescreen(NA)
                else:
                    if len(NA)==1:
                        singlescreen(NA[0],alternate_database)
                    else:
                        multiplescreen(NA,alternate_database)
        
        def overlook(BT,NAA):
            if BT.cget('text')=='Background':
                overlookscreen(NAA.background)
            elif BT.cget('text')=='Standard':
                overlookscreen(NAA.comparator)
            elif BT.cget('text')=='Sample':
                master_emissions = get_master_emissions(NAA)
                overlookscreen(NAA.sample,master_emissions)
                
        def clearall(NAA,LBS,LPKSS,LDTS,BT):
            """Clear list of selected analyte or standard spectra"""
            if BT.cget('text')=='Sample':
                NAA.sample=None
                NAA.samplemasses=None
                NAA.sampleuncertaintymasses=None
                NAA.samplemoisture=None
                NAA.sampleuncertaintymoisture=None
                LBS.configure(text='')
                LPKSS.configure(text='')
                LDTS.configure(text='')
            elif BT.cget('text')=='Standard':
                NAA.comparator=None
                NAA.comparatorcertificates=None
                NAA.comparatormasses=None
                NAA.comparatoruncertaintymasses=None
                NAA.comparatormoisture=None
                NAA.comparatoruncertaintymoisture=None
                NAA.elem_dataframe = pd.DataFrame(columns = sorted(set([el for key in NAA.certificates.keys() for el in NAA.certificates[key].keys()])))
                LBC.configure(text='')
                LPKSC.configure(text='')
                LDTC.configure(text='')
                
        def menage_comparator_masses(NAA):
            """Display single spectrum"""
            if NAA.comparator!=None:
                ratiooflines=int(rows)
                residuallenS=len(NAA.comparator)
                i=0
                paginazionespettri=[]
                paginazionemasse=[]
                paginazioneuncmasse=[]
                paginazionemoistures=[]
                paginazioneuncmoistures=[]
                while residuallenS>0:
                    try:
                        paginazionespettro=NAA.comparator[ratiooflines*i:ratiooflines*i+ratiooflines]
                        paginazionemassa=NAA.comparatormasses[ratiooflines*i:ratiooflines*i+ratiooflines]
                        paginazioneuncmassa=NAA.comparatoruncertaintymasses[ratiooflines*i:ratiooflines*i+ratiooflines]
                        paginazionemoisture=NAA.comparatormoisture[ratiooflines*i:ratiooflines*i+ratiooflines]
                        paginazioneuncmoisture=NAA.comparatoruncertaintymoisture[ratiooflines*i:ratiooflines*i+ratiooflines]
                    except IndexError:
                        paginazionespettro=NAA.comparator[ratiooflines*i:]
                        paginazionemassa=NAA.comparatormasses[ratiooflines*i:]
                        paginazioneuncmassa=NAA.comparatoruncertaintymasses[ratiooflines*i:]
                        paginazionemoisture=NAA.comparatormoisture[ratiooflines*i:]
                        paginazioneuncmoisture=NAA.comparatoruncertaintymoisture[ratiooflines*i:]
                    paginazionespettri.append(paginazionespettro)
                    paginazionemasse.append(paginazionemassa)
                    paginazioneuncmasse.append(paginazioneuncmassa)
                    paginazionemoistures.append(paginazionemoisture)
                    paginazioneuncmoistures.append(paginazioneuncmoisture)
                    residuallenS=residuallenS-ratiooflines
                    i=i+1
                SC=Toplevel()
                SC.title(NAA.comparator[0].identity+' mass management')
                SC.resizable(False,False)
                SC.focus()
                L=Label(SC).pack()
                F=Frame(SC)
                P=Frame(F)
                P.pack()
                indice=IntVar(SC)
                indice.set(0)
                massmanagementcom(P,paginazionespettri,paginazionemasse,paginazioneuncmasse,paginazionemoistures,paginazioneuncmoistures,indice,NAA)
                F.pack()
                
        def menage_comparator_emissions(NAA):
            def proceed_command(NAA,emissions_list,p_df,ETL):
                def check_selection(emission,spect):
                    try:
                        if emission in spect.assign_nuclide:
                            return True
                        else:
                            return False
                    except:
                        return False
                    
                def write_on_file(filename,NAA,emissions_list,p_df):
                    def build_buget(wb,indexsp,last_spectrum_index,item_anal_list,std_reference,r,bold,ital,Green,pct,sups,subs,sym,gray,grayit,graysub,graysym,graypct):
                        worksheet_name = f'S_{indexsp+1} {item_anal_list}'
                        wbgt = wb.add_worksheet(worksheet_name)
                        std_line = int(''.join([letter for letter in std_reference if letter.isdigit()==True]))
                        wbgt.write(0,0,'Emitter')
                        link = f'=Analysis!A{r+1}'
                        wbgt.write(0,1,link,bold)
                        wbgt.set_column('I:I', 14)
                        wbgt.set_column('L:L', 12)
                        wbgt.write(2,0,'Quantity')
                        wbgt.write(2,2,'Unit')
                        wbgt.write(2,3,'Value')
                        wbgt.write(2,4,'Std unc')
                        wbgt.write(2,5,'Rel unc')
                        wbgt.write(2,8,'Sensitivity coef.')
                        wbgt.write(2,9,'Contribution to variance')
                        wbgt.write_rich_string(3,0,ital,'X',subs,'i')
                        wbgt.write_rich_string(3,2,'[',ital,'X',subs,'i',']')
                        wbgt.write_rich_string(3,3,ital,'x',subs,'i')
                        wbgt.write_rich_string(3,4,ital,'u','(',ital,'x',subs,'i',')')
                        wbgt.write_rich_string(3,5,ital,'u',subs,'r','(',ital,'x',subs,'i',')')
                        wbgt.write_rich_string(3,6,ital,'y','(',ital,'x',subs,'i',' + ',ital,'u','(',ital,'x',subs,'i','))')
                        wbgt.write_rich_string(3,7,ital,'y','(',ital,'x',subs,'i',' - ',ital,'u','(',ital,'x',subs,'i','))')
                        wbgt.write_rich_string(3,8,ital,'c',subs,'i')
                        wbgt.write_rich_string(3,9,ital,'I',' / %')
                        wbgt.write_rich_string(4,0,ital,'C',subs,'sp std')
                        wbgt.write_rich_string(4,2,'s',sups,'-1',' g',sups,'-1')
                        link = f'=Analysis!Q{std_line}'
                        wbgt.write(4,3,link)
                        link = f'=Analysis!R{std_line}'
                        wbgt.write(4,4,link)
                        wbgt.write(5,0,'l',sym)
                        wbgt.write_rich_string(5,2,'s',sups,'-1')
                        link = f'=Analysis!C{r+1}'
                        wbgt.write(5,3,link)
                        # uncertainty on lambda still missing
                        wbgt.write_rich_string(6,0,ital,'n',subs,'p')
                        wbgt.write(6,2,'1')
                        link = f'=Analysis!B{r+1}'
                        wbgt.write(6,3,link)
                        altroindex = NAA.sample[indexsp].assign_nuclide.index(item_anal_list)
                        wbgt.write(6,4,float(NAA.sample[indexsp].peak_list[altroindex][9]))
                        wbgt.write_rich_string(7,0,ital,'t',subs,'real')
                        wbgt.write(7,2,'s')
                        link = f'=Analysis!D{last_spectrum_index+1}'
                        wbgt.write(7,3,link)
                        wbgt.write(7,4,0.1)
                        wbgt.write_rich_string(8,0,ital,'t',subs,'live')
                        wbgt.write(8,2,'s')
                        link = f'=Analysis!E{last_spectrum_index+1}'
                        wbgt.write(8,3,link)
                        wbgt.write(8,4,0.1)
                        wbgt.write_rich_string(9,0,ital,'t',subs,'decay')
                        wbgt.write(9,2,'s')
                        link = f'=Analysis!G{last_spectrum_index+1}'
                        wbgt.write(9,3,link)
                        wbgt.write(9,4,'=sqrt(120^2/3)')
                        wbgt.write_rich_string(10,0,ital,'m',subs,'sample')
                        wbgt.write(10,2,'g')
                        link = f'=Analysis!H{last_spectrum_index+1}'
                        wbgt.write(10,3,link)
                        wbgt.write(10,4,NAA.sampleuncertaintymasses[indexsp])
                        wbgt.write_rich_string(11,0,ital,'moist',subs,'sample')
                        wbgt.write(11,2,'1')
                        link = f'=Analysis!I{last_spectrum_index+1}'
                        wbgt.write(11,3,link,pct)
                        wbgt.write(11,4,NAA.sampleuncertaintymoisture[indexsp]/100,pct)
                        wbgt.write_rich_string(12,0,sym,'de',subs,'r')
                        wbgt.write_rich_string(12,2,'mm',sups,'-1')
                        wbgt.write(12,3,0.05)
                        wbgt.write(12,4,0)
                        wbgt.write_rich_string(13,0,sym,'d','d')
                        wbgt.write(13,2,'mm')
                        wbgt.write(13,3,0)
                        wbgt.write(13,4,0.2)
                        wbgt.write(15,0,'Quantity')
                        wbgt.write(15,2,'Unit')
                        wbgt.write(15,3,'Value')
                        wbgt.write(15,4,'Std unc')
                        wbgt.write(15,5,'Rel unc')
                        wbgt.write(15,9,'Contribution to variance')
                        wbgt.write(16,0,'Y',ital)
                        wbgt.write_rich_string(16,2,'[',ital,'y',']')
                        wbgt.write(16,3,'y',ital)
                        wbgt.write_rich_string(16,4,ital,'u','(',ital,'y',')')
                        wbgt.write_rich_string(16,5,ital,'u',subs,'r','(',ital,'y',')')
                        wbgt.write_rich_string(16,9,ital,'I',' / %')
                        wbgt.write_rich_string(17,0,ital,'w',subs,'a')
                        wbgt.write_rich_string(17,2,'g g',sups,'-1')
                        
                        wbgt.write(17,5,'=E18/D18',pct)
                        wbgt.write(17,9,'=sum(J5:J14)',pct)
                        
                        parm_id=['D5','D6','D7','D8','D9','D10','D11','D12','D13','D14']
                        sumofsquares =[]
                        for iiim3 in range(10):
                            fml = f'=IF(D{5+iiim3}<>0,ABS(E{5+iiim3}/D{5+iiim3}),"-")'
                            wbgt.write(4+iiim3,5,fml,pct)
                            parm_id[iiim3] = f'({parm_id[iiim3]}+{parm_id[iiim3].replace("D","E")}+1E-9)'
                            fml = f'=({parm_id[2]}*{parm_id[1]}*{parm_id[3]})/({parm_id[4]}*EXP(-{parm_id[1]}*{parm_id[5]})*(1-EXP(-{parm_id[1]}*{parm_id[3]}))*{parm_id[6]}*(1-{parm_id[7]})*{parm_id[0]})*(1-{parm_id[8]}*{parm_id[9]})'
                            wbgt.write(4+iiim3,6,fml,gray)
                            parm_id[iiim3] = f'({parm_id[iiim3]}-{parm_id[iiim3].replace("D","E")}-1E-9)'
                            fml = f'=({parm_id[2]}*{parm_id[1]}*{parm_id[3]})/({parm_id[4]}*EXP(-{parm_id[1]}*{parm_id[5]})*(1-EXP(-{parm_id[1]}*{parm_id[3]}))*{parm_id[6]}*(1-{parm_id[7]})*{parm_id[0]})*(1-{parm_id[8]}*{parm_id[9]})'
                            wbgt.write(4+iiim3,7,fml,gray)
                            parm_id=['D5','D6','D7','D8','D9','D10','D11','D12','D13','D14']
                            fml = f'=(G{5+iiim3}-H{5+iiim3})/(2*E{5+iiim3}+2E-9)'
                            wbgt.write(4+iiim3,8,fml)
                            fml = f'=(E{5+iiim3}*I{5+iiim3})^2/E18^2'
                            wbgt.write(4+iiim3,9,fml,pct)
                            sumofsquares.append(f'(E{5+iiim3}*I{5+iiim3})^2')
                        
                        sumofs = '+'.join(sumofsquares)
                        sumofs = f'=sqrt({sumofs})'
                        fml = f'=({parm_id[2]}*{parm_id[1]}*{parm_id[3]})/({parm_id[4]}*EXP(-{parm_id[1]}*{parm_id[5]})*(1-EXP(-{parm_id[1]}*{parm_id[3]}))*{parm_id[6]}*(1-{parm_id[7]})*{parm_id[0]})*(1-{parm_id[8]}*{parm_id[9]})'
                        wbgt.write(17,3,fml,Green)
                        wbgt.write(17,4,sumofs,Green)
                        wbgt.set_column('G:H', None, None, {'hidden': True})
                        
                        link1,link2 = f"='{worksheet_name}'!D18",f"='{worksheet_name}'!E18"
                        return link1,link2
                    
                    def graph_sheet(ws,dict_spc,bold,est):
                        def get_target(key):
                            elem, energy = key.split()
                            for item in A:
                                if str(item[2]) == energy and item[1] == elem:
                                    return item[0]
                        x_size = 10
                        y_size = 15
                        tp_clmn = (('B','C'),('D','E'),('F','G'),('H','I'),('J','K'),('L','M'),('N','O')) 
                        every_target = tuple(sorted(set([get_target(k) for i in dict_spc for k in dict_spc[i].keys()])))
                        every_isot = tuple([tuple(sorted(set([k for i in dict_spc for k in dict_spc[i] if get_target(k)==l]))) for l in every_target])
                        r = 0
                        ws.write(0,1,'x / g g-1')
                        ws.write(0,2,'U(x) / g g-1, k=2')
                        for i,ttg in zip(every_isot,every_target):
                            chart_row = int(r)
                            ws.write(r,0,ttg,bold)
                            chrt = wb.add_chart({'type':'column'})
                            r += 1
                            ll = 0
                            for s,column_l in zip(i,tp_clmn):
                                ws.write(r,1+ll,s.split()[-1],bold)
                                ll += 2
                                chrt.add_series({'name': f'{s} keV', 'values': f'=Graphs!{column_l[0]}{r+2}:{column_l[0]}{r+1+len(dict_spc)}', 'categories': f'=Graphs!A{r+2}:A{r+1+len(dict_spc)}', 'y_error_bars': {'type' : 'custom', 'plus_values': f'=Graphs!{column_l[1]}{r+2}:{column_l[1]}{r+1+len(dict_spc)}', 'minus_values': f'=Graphs!{column_l[1]}{r+2}:{column_l[1]}{r+1+len(dict_spc)}'}})
                            for k in dict_spc:
                                r += 1
                                ll = 0
                                for s in i:
                                    line = dict_spc[k].get(s,'')
                                    if line != '':
                                        line1, line2 = line[0], line[1]
                                    else:
                                        line1, line2 = '', ''
                                    if ll==0:
                                        ws.write(r,0,k,est)
                                    ws.write(r,1+ll,line1)
                                    ws.write(r,2+ll,line2)
                                    ll += 2
                            chrt.set_y_axis({'name': 'w / g g-1', 'name_font': {'size': 11}, 'num_format': '0.00E+00'})
                            chrt.set_title({'name': f'{ttg}'})
                            chrt.set_size({'width': x_size*64, 'height': y_size*20})
                            ws.insert_chart(chart_row, len(i)*2+2, chrt)
                            if len(dict_spc)+3 < 15:
                                r = r + (16-r%16)
                            else:
                                r += 2

                    def search_lambda_from_database(emission):
                        isot, energ = emission.split()
                        for i in A:
                            if i[1] == isot and str(i[2]) == energ:
                                tm = i[3]
                                break
                        return np.log(2)/tm
                    wb = xlsxwriter.Workbook(filename)
                    ws = wb.add_worksheet('Analysis')
                    wbwg = wb.add_worksheet('Graphs')
                    bold = wb.add_format({'bold': True})
                    ital = wb.add_format({'italic': True})
                    Green = wb.add_format({'bold': True, 'font_color': 'green', 'num_format':'0.00E+00'})
                    Greenu = wb.add_format({'bold': True, 'font_color': 'green', 'num_format':'0.0E+00'})
                    Greenppm = wb.add_format({'bold': True, 'font_color': 'green', 'num_format':'0.000'})
                    pct = wb.add_format({'num_format': 0x0a})
                    sups = wb.add_format({'font_script': 1})
                    subs = wb.add_format({'font_script': 2})
                    gray = wb.add_format({'font_color': 'gray'})
                    grayit = wb.add_format({'italic': True, 'font_color': 'gray'})
                    graysub = wb.add_format({'font_script': 2, 'font_color': 'gray'})
                    graypct = wb.add_format({'num_format': 0x0a, 'font_color': 'gray'})
                    est = wb.add_format({'align': 'right'})
                    DL_red = wb.add_format({'font_color': 'red', 'italic': True, 'num_format':'0.0E+00'})
                    DL_redppm = wb.add_format({'font_color': 'red', 'italic': True, 'num_format':'0.000'})
                    dateandtime = wb.add_format({'num_format': 'dd/mm/yyyy hh:mm'})
                    try:
                        sym = wb.add_format({'font_name': 'Symbol'})
                        graysym = wb.add_format({'font_name': 'Symbol', 'font_color': 'gray'})
                    except:
                        sym = wb.add_format({'font_name': 'Times New Roman'})
                        graysym = wb.add_format({'font_name': 'Times New Roman', 'font_color': 'gray'})
                    check='non-1/v reaction'
                    emission_standard_list = {}
                    ws.set_column('A:C', 16)
                    ws.set_column('G:L', 12)
                    ws.write(0,0,'Irradiation end')
                    ws.write_rich_string(0,2,ital,'t',subs,'i',' / s')
                    ws.write(0,4,'Channel')
                    ws.write_rich_string(0,5,ital,'f',' / 1')
                    ws.write_rich_string(0,6,sym,'a',' / 1')
                    ws.write(0,8,'Geometry')
                    ws.write(0,9,'Distance')
                    ws.write(0,10,'Detector')
                    ws.write(0,12,f'Relative-LENA version {NAA.info["version"]}')
                    ws.write(1,12,f'CRM database: {NAA.info["version_CRM"]}')
                    ws.write(2,12,f'Emission database: {NAA.info["version_emissions"]}')
                    ws.write(1,0,NAA.irradiation.datetime,dateandtime)
                    ws.write(1,2,NAA.irradiation.time)
                    ws.write(1,4,NAA.irradiation.channel)
                    ws.write(1,5,NAA.irradiation.f)
                    ws.write(1,6,NAA.irradiation.a)
                    ws.write(1,8,NAA.experimental_additional_information.get('geometry',''))
                    ws.write(1,9,NAA.experimental_additional_information.get('distance',''))
                    ws.write(1,10,NAA.experimental_additional_information.get('detector',''))
                    ws.write(3,0,'Standards',bold)
                    ws.write(4,0,'Emission')
                    ws.write(4,1,'Spectrum')
                    ws.write(4,2,'Acquisition date')
                    ws.write_rich_string(4,3,ital,'t',subs,'real',' / s')
                    ws.write_rich_string(4,4,ital,'t',subs,'live',' / s')
                    ws.write_rich_string(4,5,ital,'t',subs,'dead',' / 1')
                    ws.write_rich_string(4,6,ital,'t',subs,'decay',' / s')
                    ws.write_rich_string(4,7,ital,'n',subs,'p',' / 1')
                    ws.write_rich_string(4,8,sym,'l',' / s',sups,'-1')
                    ws.write(4,9,'mass / g')
                    ws.write(4,10,'moisture / 1')
                    ws.write_rich_string(4,11,ital,'w',' / µg g',sups,'-1')
                    ws.write(4,12,'Certificate')
                    ws.write_rich_string(4,13,'Sp. c. rate / s',sups,'-1',' g',sups,'-1')
                    ws.write_rich_string(4,14,ital,'u',subs,'r','(Sp. c. rate) / 1')
                    ws.write_rich_string(4,16,ital,'w.','Aver. / s',sups,'-1',' g',sups,'-1')
                    ws.write_rich_string(4,17,ital,'u','(',ital,'w.','Aver.) / s',sups,'-1',' g',sups,'-1')
                    ws.write(4,18,'Rel std deviation')
                    r = 5
                    for iii,eee in zip(emissions_list,NAA.emission_list):
                        little_selection = [check_selection(iii,spect) for spect in NAA.comparator]
                        target = iii.split()[0]
                        for targetfromA in A:
                            if targetfromA[1] == target:
                                target = targetfromA[0]
                                break
                        pls_df = p_df[little_selection]
                        pls_df = pls_df[target]
                        ws.write(r,0,iii,bold)
                        for iline_index,iline in enumerate(pls_df[eee]):
                            for spectra_info in NAA.comparator:
                                if spectra_info.spectrumpath == pls_df[eee].index[iline_index]:
                                    spectrum_info = spectra_info
                                    break
                            ws.write(r+iline_index,1,spectrum_info.filename(),est)
                            ws.write(r+iline_index,2,spectrum_info.datetime,dateandtime)
                            ws.write(r+iline_index,3,spectrum_info.real_time)
                            ws.write(r+iline_index,4,spectrum_info.live_time)
                            fml = f'=(D{r+iline_index+1}-E{r+iline_index+1})/D{r+iline_index+1}'
                            ws.write(r+iline_index,5,fml,pct)
                            fml = f'=(C{r+iline_index+1}-A2)*86400'
                            ws.write(r+iline_index,6,fml)
                            iii_index = spectrum_info.assign_nuclide.index(iii)
                            ws.write(r+iline_index,7,float(spectrum_info.peak_list[iii_index][8]))
                            ws.write(r+iline_index,8,search_lambda_from_database(iii))
                            idspcnfo = NAA.comparator.index(spectrum_info)
                            ws.write(r+iline_index,9,NAA.comparatormasses[idspcnfo])
                            ws.write(r+iline_index,10,NAA.comparatormoisture[idspcnfo]/100,pct)
                            if np.isnan(iline):
                                iline = 0.0
                            ws.write(r+iline_index,11,iline)
                            try:
                                ws.write(r+iline_index,12,spectrum_info.selected_certificate)
                                literature_uncertainty = NAA.certificates[spectrum_info.selected_certificate][target][1]
                            except AttributeError:
                                literature_uncertainty = 0
                            fml = f'=(H{r+iline_index+1}*I{r+iline_index+1}*D{r+iline_index+1})/(E{r+iline_index+1}*EXP(-I{r+iline_index+1}*G{r+iline_index+1})*(1-EXP(-I{r+iline_index+1}*D{r+iline_index+1}))*J{r+iline_index+1}*(1-K{r+iline_index+1})*L{r+iline_index+1}/1E6)'
                            ws.write(r+iline_index,13,fml)
                            fml = f'=sqrt(({float(spectrum_info.peak_list[iii_index][9])}/H{r+iline_index+1})^2+({NAA.comparatoruncertaintymasses[idspcnfo]}/J{r+iline_index+1})^2+({literature_uncertainty}/L{r+iline_index+1})^2+{NAA.comparatoruncertaintymoisture[idspcnfo]/100}^2)'
                            ws.write(r+iline_index,14,fml,pct)
                            if iline_index == 0:
                                #numm,denm = [f'N{r+iline_index+1+ineee}^2*O{r+iline_index+1+ineee}' for ineee in range(len(pls_df[eee]))], [f'N{r+iline_index+1+ineee}*O{r+iline_index+1+ineee}' for ineee in range(len(pls_df[eee]))]
                                numm,denm = [f'N{r+iline_index+1+ineee}^3*O{r+iline_index+1+ineee}^2' for ineee in range(len(pls_df[eee]))], [f'N{r+iline_index+1+ineee}*O{r+iline_index+1+ineee}' for ineee in range(len(pls_df[eee]))]
                                summ = ')^2+('.join(denm)
                                numm, denm = '+'.join(numm), '+'.join(denm)
                                #fml = f'=({numm})/({denm})'
                                fml = f'=({numm})/(({summ})^2)'
                                ws.write(r+iline_index,16,fml,Green)
                                fml = f'=IFERROR(1/(sqrt({len(pls_df[eee])}/(({summ})^2))),O{r+iline_index+1}*Q{r+iline_index+1})'
                                ws.write(r+iline_index,17,fml,Green)
                                emission_standard_list[iii] = f'Q{r+iline_index+1}'
                                fml = f'=R{r+iline_index+1}/Q{r+iline_index+1}'
                                ws.write(r+iline_index,18,fml,pct)
                        if len(pls_df[eee])!=0:
                            r += len(pls_df[eee])+1
                        else:
                            ws.write(r,13,'-')
                            ws.write(r,14,'-')
                            fml = f'=average(N{r+1}:N{r+1})'
                            ws.write(r,16,fml,Green)
                            fml = f'=O{r+1}*Q{r+1}'
                            ws.write(r,17,fml,Green)
                            emission_standard_list[iii] = f'Q{r+1}'
                            r += 2
                    ws.write(r,0,'Analytes',bold)
                    r += 1
                    dict_of_acquired_spectra = {}
                    for anal_spect in NAA.sample:
                        index_massesofsamples = NAA.sample.index(anal_spect)
                        if anal_spect.assign_nuclide == None:
                            anal_spect.assign_nuclide=['']*len(anal_spect.peak_list)
                            master_emissions = get_master_emissions(NAA)
                            funtion_autoselect_emission(anal_spect,master_emissions)
                        mini_anal_list = [x_assign_nuclide for x_assign_nuclide in anal_spect.assign_nuclide if x_assign_nuclide!='']
                        if len(mini_anal_list)!=0:
                            mini_anal_list.sort(key=lambda item : (item.split()[0],item.split()[-1]))
                        ws.write(r,1,'Spectrum')
                        ws.write(r,2,'Acquisition date')
                        ws.write_rich_string(r,3,ital,'t',subs,'real',' / s')
                        ws.write_rich_string(r,4,ital,'t',subs,'live',' / s')
                        ws.write_rich_string(r,5,ital,'t',subs,'dead',' / 1')
                        ws.write_rich_string(r,6,ital,'t',subs,'decay',' / s')
                        ws.write(r,7,'mass / g')
                        ws.write(r,8,'moisture / 1')
                        r += 1
                        last_spectrum_index = int(r)
                        ws.write(r,1,anal_spect.filename(),est) #spectrumpath
                        ws.write(r,2,anal_spect.datetime,dateandtime)
                        ws.write(r,3,anal_spect.real_time)
                        ws.write(r,4,anal_spect.live_time)
                        fml = f'=(D{r+1}-E{r+1})/D{r+1}'
                        ws.write(r,5,fml,pct)
                        fml = f'=(C{r+1}-A2)*86400'
                        ws.write(r,6,fml)
                        ws.write(r,7,NAA.samplemasses[index_massesofsamples])
                        ws.write(r,8,NAA.samplemoisture[index_massesofsamples]/100,pct)
                        r += 2
                        ws.write(r,0,'Emission')
                        ws.write_rich_string(r,1,ital,'n',subs,'p',' / 1')
                        ws.write_rich_string(r,2,sym,'l',' / s',sups,'-1')
                        ws.write_rich_string(r,4,ital,'w',subs,'a',' / g g',sups,'-1')
                        ws.write_rich_string(r,5,ital,'u','(',ital,'w',subs,'a',') / g g',sups,'-1')
                        ws.write_rich_string(r,6,ital,'u',subs,'r','(',ital,'w',subs,'a',') / g g',sups,'-1')
                        ws.write_rich_string(r,7,'Detection limit / g g',sups,'-1')
                        ws.write_rich_string(r,9,ital,'w',' / µg g',sups,'-1')
                        ws.write_rich_string(r,10,ital,'u','(',ital,'w',') / µg g',sups,'-1')
                        ws.write_rich_string(r,11,'Detection limit / µg g',sups,'-1')
                        items_in_spectra = {}
                        r += 1
                        for item_anal_list in mini_anal_list:
                            keyname = f'{item_anal_list.split()[0].split("-")[0]} {item_anal_list.split()[-1]}'
                            items_in_spectra[item_anal_list] = (f'=Analysis!E{r+1}',f'=2*Analysis!F{r+1}')
                            #items_in_spectra[keyname] = (f'=Analysis!E{r+1}',f'=2*Analysis!F{r+1}')
                            ws.write(r,0,item_anal_list,bold)
                            ass_nul_index_g = anal_spect.assign_nuclide.index(item_anal_list)
                            ws.write(r,1,float(anal_spect.peak_list[ass_nul_index_g][8]))
                            ws.write(r,2,search_lambda_from_database(item_anal_list))
                            link1,link2 = build_buget(wb,index_massesofsamples,last_spectrum_index,item_anal_list,emission_standard_list.get(item_anal_list,0),r,bold,ital,Green,pct,sups,subs,sym,gray,grayit,graysub,graysym,graypct)
                            #ws.write(r,4,link1,Green)#wa
                            fml = f'=IF(B{r+1}>1,{link1.replace("=","")},"not quantified")'
                            ws.write(r,4,fml,Green)#here
                            fml = f'=IF(B{r+1}>1,{link2.replace("=","")},"-")'
                            ws.write(r,5,fml,Greenu)#uwa
                            fml = f'=F{r+1}/E{r+1}'
                            ws.write(r,6,fml,pct)#urwa
                            fml = f'=((2.71+4.65*sqrt({float(anal_spect.peak_list[ass_nul_index_g][20])}))*C{r+1}*D{last_spectrum_index+1})/(E{last_spectrum_index+1}*EXP(-C{r+1}*G{last_spectrum_index+1})*(1-EXP(-C{r+1}*D{last_spectrum_index+1}))*H{last_spectrum_index+1}*(1-I{last_spectrum_index+1})*{emission_standard_list.get(item_anal_list,0)})'
                            ws.write(r,7,fml,DL_red)#det_lim
                            fml = f'=E{r+1}*1E6'
                            ws.write(r,9,fml,Greenppm)
                            fml = f'=F{r+1}*1E6'
                            ws.write(r,10,fml,Greenppm)
                            fml = f'=H{r+1}*1E6'
                            ws.write(r,11,fml,DL_redppm)
                            r += 1
                        if len(mini_anal_list)==0:
                            ws.write(r,0,'-')
                            r += 1
                        r += 1
                        dict_of_acquired_spectra[f'=Analysis!B{last_spectrum_index+1}'] = items_in_spectra
                    graph_sheet(wbwg,dict_of_acquired_spectra,bold,est)
                    wb.close()
                
                cdn=ETL.winfo_children()
                for i in cdn:
                    i.destroy()
                F=Frame(ETL)
                LMN=Label(F, text='\nprocessing...\nit may take a while,\nhave a cup of something\n')
                LMN.pack(padx=30)
                F.pack()
                types=[('Excel file','.xlsx')]
                filename=asksaveasfilename(filetypes=types, defaultextension='.xlsx')
                if filename!='' and filename!=None:
                    write_on_file(filename,NAA,emissions_list,p_df)
                    ETL.title('Processing complete')
                    LMN.configure(text=f'\n{os.path.basename(filename)} saved!\n')
                    ETL.focus()
                    event="<FocusOut>"
                    ETL.bind(event,lambda event=event,M=ETL : M.destroy())

            def find_elements():
                something_beautiful = []
                if NAA.comparator is not None:
                    for spec in NAA.comparator:
                        if spec.assign_nuclide != None:
                            something_beautiful += [nuc.split()[0] for nuc in spec.assign_nuclide if nuc!='']
                something_beautiful = sorted(set(something_beautiful))
                sel = []
                for e_comp in something_beautiful:
                    for item in A:
                        if item[1] == e_comp:
                            sel.append(item[0])
                            break
                sel = sorted(set(sel))
                return sel
            
            def std_cert(spectrum):
                try:
                    NAA.certificates.get(spectrum.selected_certificate,'')
                except AttributeError:
                    return ''
                else:
                    return spectrum.selected_certificate
                
            def indice_pagina_plus(PP,paginazioneemissioni,indice_pagina,paginazionetrues,trgts,listofresults,emi_list,L_mean_sd):
                if indice_pagina.get()<len(paginazioneemissioni)-1:
                    indice_pagina.set(indice_pagina.get()+1)
                    sciogliipicchianaP(PP,paginazioneemissioni,indice_pagina,paginazionetrues,trgts,listofresults,emi_list,L_mean_sd)
            
            def indice_pagina_minus(PP,paginazioneemissioni,indice_pagina,paginazionetrues,trgts,listofresults,emi_list,L_mean_sd):
                if indice_pagina.get()>0:
                    indice_pagina.set(indice_pagina.get()-1)
                    sciogliipicchianaP(PP,paginazioneemissioni,indice_pagina,paginazionetrues,trgts,listofresults,emi_list,L_mean_sd)
                    
            def command_check_CBT(CB,paginazionetrues,indice,ii,var,emi_list,LR,LSD):
                def average_preview(LR,LT,LB):
                    sm = np.array([lr for lr,lt in zip(LR,LT) if lt==True])
                    LB.configure(text=f'ȳ ={format(sm.mean(),".3e")},     rel. σ ={round(sm.std()/sm.mean()*100,2)} %')
                if paginazionetrues[indice.get()][ii] == True:
                    CB.deselect()
                    paginazionetrues[indice.get()][ii] = False
                    emi_list[indice.get()*int(rows)+ii] = False
                    average_preview(LR,emi_list,LSD)
                else:
                    CB.select()
                    paginazionetrues[indice.get()][ii] = True
                    emi_list[indice.get()*int(rows)+ii] = True
                    average_preview(LR,emi_list,LSD)
                
            def sciogliipicchianaP(PP,paginazioneemissioni,indice_pagina,paginazionetrues,trgts,listofresults,emi_list,L_mean_sd):
                if indice_pagina.get()>-1 and indice_pagina.get()<len(paginazioneemissioni):
                    cdn=PP.winfo_children()
                    for i in cdn:
                        i.destroy()
                    F=Frame(PP)
                    chrw=12
                    Bmeno=Button(F, text='<', relief='flat', command=lambda PP=PP,paginazioneemissioni=paginazioneemissioni,indice_pagina=indice_pagina,paginazionetrues=paginazionetrues,trgts=trgts,listofresults=listofresults : indice_pagina_minus(PP,paginazioneemissioni,indice_pagina,paginazionetrues,trgts,listofresults,emi_list,L_mean_sd)).pack(side=LEFT)
                    L=Label(F, text=f'page {indice_pagina.get()+1} of {len(paginazioneemissioni)}').pack(side=LEFT)
                    Bpiu=Button(F, text='>', relief='flat', command=lambda PP=PP,paginazioneemissioni=paginazioneemissioni,indice_pagina=indice_pagina,paginazionetrues=paginazionetrues,trgts=trgts,listofresults=listofresults : indice_pagina_plus(PP,paginazioneemissioni,indice_pagina,paginazionetrues,trgts,listofresults,emi_list,L_mean_sd)).pack(side=LEFT)
                    F.pack()
                    titles=['filename',trgts,'selection','result','certificate','peak unc.']
                    FL=Frame(PP)
                    L=Label(FL, text=titles[2], width=chrw-3).pack(side=LEFT)
                    L=Label(FL, text=titles[0], width=chrw*5).pack(side=LEFT)
                    L=Label(FL, text=titles[4], width=chrw).pack(side=LEFT)
                    L=Label(FL, text=titles[1], width=chrw).pack(side=LEFT)
                    L=Label(FL, text=titles[5], width=chrw).pack(side=LEFT)
                    L=Label(FL, text=titles[3], width=chrw).pack(side=LEFT)
                    FL.pack(anchor=W, padx=2)
                    valuei=0
                    for ii,j in enumerate(paginazioneemissioni[indice_pagina.get()].index):
                        FL=Frame(PP)
                        varCBT=StringVar(F)
                        CBT=Checkbutton(FL, text='', variable=varCBT, onvalue='True', offvalue='False', width=6)
                        CBT.pack(side=LEFT)
                        if paginazionetrues[indice_pagina.get()][ii] == True:
                            CBT.select()
                        else:
                            CBT.deselect()
                        CBT.configure(command=lambda CB=CBT,paginazionetrues=paginazionetrues,indice=indice_pagina,ii=ii,var=varCBT,emi_list=emi_list,LR=listofresults,LSD=L_mean_sd : command_check_CBT(CB,paginazionetrues,indice,ii,var,emi_list,LR,LSD))
                        L=Label(FL, text=str(j.split('/')[-1]), width=chrw*5, anchor=E).pack(side=LEFT) #filename
                        cert_p = ''
                        for spectrum_p in NAA.comparator:
                            if spectrum_p.spectrumpath == j:
                                try:
                                    cert_p = spectrum_p.selected_certificate
                                except AttributeError:
                                    cert_p = ''
                                break
                        L=Label(FL, text=cert_p, width=chrw).pack(side=LEFT)
                        L=Label(FL, text=str(float(paginazioneemissioni[indice_pagina.get()].loc[j]).__round__(2)), width=chrw).pack(side=LEFT) #certificate value
                        for spectrum_path_index in NAA.comparator:
                            if spectrum_path_index.spectrumpath == j:
                                ddx = spectrum_path_index.assign_nuclide.index(emissions_list[indice.get()])
                                uncddx, valddx = float(spectrum_path_index.peak_list[ddx][9]),float(spectrum_path_index.peak_list[ddx][8])
                        L=Label(FL, text=f'{round(uncddx/valddx*100,2)} %', width=chrw).pack(side=LEFT)
                        L=Label(FL, text=f'{format(listofresults[indice_pagina.get()*int(rows)+ii],".3e")}', width=chrw).pack(side=LEFT)
                        FL.pack(anchor=W, padx=2)
                
            def sciorina_emissions(P,indice,emissions_list,NAA,partial_df):
                def search_lambda_from_database(emission):
                    isot, energ = emission.split()
                    for i in A:
                        if i[1] == isot and str(i[2]) == energ:
                            tm = i[3]
                            break
                    return np.log(2)/tm
                
                def result_preview(df,emission,spectrumname,NAA,lamd):
                    try:
                        ppm = df.loc[spectrumname]
                        for i in NAA.comparator:
                            if i.spectrumpath == spectrumname:
                                neta = float(i.peak_list[i.assign_nuclide.index(emission)][8])
                                rt, lt = i.real_time, i.live_time
                                dt = i.datetime-NAA.irradiation.datetime
                                dt = dt.days*86400+dt.seconds
                                idx = NAA.comparator.index(i)
                                break
                        mm, mois = NAA.comparatormasses[idx], NAA.comparatormoisture[idx]
                        w = mm * (1-mois/100) * ppm * 1E-6
                        return (neta*lamd*rt)/(lt*np.exp(-lamd*dt)*(1-np.exp(-lamd*rt))*w)
                    except:
                        return np.nan
                    
                def average_preview(LR,LT,LB):
                    sm = np.array([lr for lr,lt in zip(LR,LT) if lt==True])
                    LB.configure(text=f'ȳ ={format(sm.mean(),".3e")},     rel. σ ={round(sm.std()/sm.mean()*100,2)} %')
                
                def check_selection(emission,spect):
                    try:
                        if emission in spect.assign_nuclide:
                            return True
                        else:
                            return False
                    except:
                        return False
                
                FLE = Frame(P)
                L=Label(FLE, text=f'{emissions_list[indice.get()]} keV', anchor=W).pack(padx=2, anchor=W, side=LEFT)
                L=Label(FLE, text='', width=6).pack(side=LEFT)
                L_mean_sd = Label(FLE, text=f'ȳ ={0.0},     σ ={0.0}', anchor=W)
                L_mean_sd.pack(padx=2, anchor=W, side=LEFT)
                FLE.pack(anchor=W, expand=1, fill=X)
                for item in A:
                    if item[1] == emissions_list[indice.get()].split()[0]:
                        emitr = item[0]
                        break
                little_selection = [check_selection(emissions_list[indice.get()],spect) for spect in NAA.comparator]
                i = 0
                residuallenS=len(partial_df.loc[little_selection,emitr])
                ratiooflines=int(rows)
                lamd = search_lambda_from_database(emissions_list[indice.get()])
                listofresults = [result_preview(partial_df.loc[little_selection,emitr],emissions_list[indice.get()],spectrumname,NAA,lamd) for spectrumname in partial_df.loc[little_selection,emitr].index]
                if NAA.emission_list[indice.get()] is not None:
                    listoftrues = NAA.emission_list[indice.get()]
                else:
                    listoftrues = [True]*residuallenS
                    NAA.emission_list[indice.get()] = listoftrues
                paginazioneemissioni = []
                paginazionetrues = []
                average_preview(listofresults,listoftrues,L_mean_sd)
                while residuallenS>0:
                    try:
                        paginaemi=partial_df.loc[little_selection,emitr][ratiooflines*i:ratiooflines*i+ratiooflines]
                        paginazionetrue=listoftrues[ratiooflines*i:ratiooflines*i+ratiooflines]
                    except IndexError:
                        paginaemi=partial_df.loc[little_selection,emitr][ratiooflines*i:]
                        paginazionetrue=listoftrues[ratiooflines*i:]
                    paginazioneemissioni.append(paginaemi)
                    paginazionetrues.append(paginazionetrue)
                    residuallenS=residuallenS-ratiooflines
                    i=i+1
                indice_pagina = IntVar(P)
                indice_pagina.set(0)
                PP = Frame(P)
                sciogliipicchianaP(PP,paginazioneemissioni,indice_pagina,paginazionetrues,emitr,listofresults,NAA.emission_list[indice.get()],L_mean_sd)
                PP.pack(anchor=W)
                
            def emission_plus(indice,ETL,emissions_list,NAA,partial_df):
                if indice.get()<len(emissions_list)-1:
                    indice.set(indice.get()+1)
                    cdn=ETL.winfo_children()
                    for i in cdn:
                        i.destroy()
                    F=Frame(ETL)
                    chrw=10
                    Bmeno=Button(F, text='<', relief='flat', command=lambda indice=indice,ETL=ETL,emissions_list=emissions_list,NAA=NAA,partial_df=partial_df : emission_minus(indice,ETL,emissions_list,NAA,partial_df)).pack(side=LEFT)
                    L=Label(F, text=f'emission {indice.get()+1} of {len(emissions_list)}').pack(side=LEFT)
                    Bpiu=Button(F, text='>', relief='flat', command=lambda indice=indice,ETL=ETL,emissions_list=emissions_list,NAA=NAA,partial_df=partial_df : emission_plus(indice,ETL,emissions_list,NAA,partial_df)).pack(side=LEFT)
                    F.pack()
                    if indice.get()==len(emissions_list)-1:
                        BCTN=Button(ETL,text='Proceed',width=10, command=lambda NAA=NAA,emissions_list=emissions_list,p_df=partial_df,ETL=ETL : proceed_command(NAA,emissions_list,p_df,ETL)).pack()
                    else:
                        L=Label(ETL).pack()
                    P = Frame(ETL)
                    sciorina_emissions(P,indice,emissions_list,NAA,partial_df)
                    P.pack(anchor=W)
            
            def emission_minus(indice,ETL,emissions_list,NAA,partial_df):
                if indice.get()>0:
                    indice.set(indice.get()-1)
                    cdn=ETL.winfo_children()
                    for i in cdn:
                        i.destroy()
                    F=Frame(ETL)
                    chrw=10
                    Bmeno=Button(F, text='<', relief='flat', command=lambda indice=indice,ETL=ETL,emissions_list=emissions_list,NAA=NAA,partial_df=partial_df : emission_minus(indice,ETL,emissions_list,NAA,partial_df)).pack(side=LEFT)
                    L=Label(F, text=f'emission {indice.get()+1} of {len(emissions_list)}').pack(side=LEFT)
                    Bpiu=Button(F, text='>', relief='flat', command=lambda indice=indice,ETL=ETL,emissions_list=emissions_list,NAA=NAA,partial_df=partial_df : emission_plus(indice,ETL,emissions_list,NAA,partial_df)).pack(side=LEFT)
                    F.pack()
                    L=Label(ETL).pack()
                    P = Frame(ETL)
                    sciorina_emissions(P,indice,emissions_list,NAA,partial_df)
                    P.pack(anchor=W)
                
            def emission_peaks(spec):
                def search_in_A(temp):
                    final_target = []
                    for i in temp:
                        emis,energy = i.split()
                        for line in A:
                            if emis==line[1] and energy==str(line[2]):
                                final_target.append(line)
                    short_final_target = sorted(set([i[0] for i in final_target]))
                    return final_target,short_final_target
                if spec.assign_nuclide!=None:
                    temp = sorted(set(spec.assign_nuclide))
                    if '' in temp:
                        temp.remove('')
                    final_target,short_final_target = search_in_A(temp)
                    return  final_target,short_final_target
                else:
                    return [],[]
                    
            multiple_selection = False
            try:
                for anal_spect in NAA.sample:
                    if anal_spect.assign_nuclide == None:
                        anal_spect.assign_nuclide=['']*len(anal_spect.peak_list)
                        master_emissions = get_master_emissions(NAA)
                        funtion_autoselect_emission(anal_spect,master_emissions)
                    if len([item for item in anal_spect.assign_nuclide if item!='']) != len(set([item for item in anal_spect.assign_nuclide if item!=''])):
                        multiple_selection = True
                        break
            except TypeError:
                pass
            if NAA.comparator!=None and NAA.irradiation!=None and NAA.sample!=None and 0 not in NAA.comparatormasses and 0 not in NAA.samplemasses and multiple_selection==False:
                emissions_list = []
                for spect in NAA.comparator:
                    try:
                        emissions_list += [emiss for emiss in spect.assign_nuclide if emiss!='']
                    except TypeError:
                        pass
                emissions_list = list(set(emissions_list))
                emissions_list.sort(key = lambda item : (item.split()[0],float(item.split()[-1])), reverse=True)
                NAA.comparatorcertificates=[std_cert(spectrum) for spectrum in NAA.comparator]
                NAA.emission_list = [None] * len(emissions_list)
                emissions_in_spectra = []
                element_selection = find_elements()
                for spec,certf in zip(NAA.comparator,NAA.comparatorcertificates):
                    long_t, short_t = emission_peaks(spec)
                    emissions_in_spectra.append([short_t, long_t])
                    if certf!='':
                        NAA.elem_dataframe.loc[spec.spectrumpath] = {key:value[0] for (key,value) in NAA.certificates[certf].items() if key in element_selection}
                partial_df = NAA.elem_dataframe[element_selection]
                if len(emissions_list)>0:
                    ETL = Toplevel()
                    ETL.resizable(False,False)
                    ETL.title('Standard emissions average')
                    ETL.focus()
                    indice=IntVar(ETL)
                    indice.set(0)
                    F=Frame(ETL)
                    chrw=10
                    Bmeno=Button(F, text='<', relief='flat', command=lambda indice=indice,ETL=ETL,emissions_list=emissions_list,NAA=NAA,partial_df=partial_df : emission_minus(indice,ETL,emissions_list,NAA,partial_df)).pack(side=LEFT)
                    L=Label(F, text=f'emission {indice.get()+1} of {len(emissions_list)}').pack(side=LEFT)
                    Bpiu=Button(F, text='>', relief='flat', command=lambda indice=indice,ETL=ETL,emissions_list=emissions_list,NAA=NAA,partial_df=partial_df : emission_plus(indice,ETL,emissions_list,NAA,partial_df)).pack(side=LEFT)
                    F.pack()
                    if indice.get()==len(emissions_list)-1:
                        BCTN=Button(ETL,text='Proceed',width=10, command=lambda NAA=NAA,emissions_list=emissions_list,p_df=partial_df,ETL=ETL : proceed_command(NAA,emissions_list,p_df,ETL)).pack()
                    else:
                        L=Label(ETL).pack()
                    P = Frame(ETL)
                    sciorina_emissions(P,indice,emissions_list,NAA,partial_df)
                    P.pack(anchor=W)
                else:
                    message_expl = '\nNo emissions selected from standard spectra\n'
                    messagebox.showwarning('Missing data', message_expl)
            else:
                message_expl = '\nGeneric Error!\nCheck somewhere in the window\n'
                if NAA.irradiation==None:
                    message_expl = '\nIrradiation is not selected\n'
                elif NAA.comparator==None:
                    message_expl = '\nStandards are missing\n'
                elif 0 in NAA.comparatormasses:
                    message_expl = '\nMass of one or multiple standards is 0\n'
                elif NAA.sample==None:
                    message_expl = '\nSamples are missing\n'
                elif 0 in NAA.samplemasses:
                    message_expl = '\nMass of one or multiple samples is 0\n'
                elif multiple_selection==True:
                    message_expl = '\nMultiple peaks assigned to the same emission\nmanually check spectra or decrease energy tolerance\n'
                messagebox.showwarning('Missing data', message_expl)
                        
        def f_mass(mass,moisture):
            try:
                 return float(mass.get())*(1-float(moisture.get())/100)
            except:
                 return 0.0
             
        def masses_plus(P,paginazionespettri,paginazionemasse,paginazioneuncmasse,paginazionemoistures,paginazioneuncmoistures,indice,NAA,typ='standard'):
            if indice.get()<len(paginazionespettri)-1:
                indice.set(indice.get()+1)
                massmanagementcom(P,paginazionespettri,paginazionemasse,paginazioneuncmasse,paginazionemoistures,paginazioneuncmoistures,indice,NAA,typ)
            
        def masses_minus(P,paginazionespettri,paginazionemasse,paginazioneuncmasse,paginazionemoistures,paginazioneuncmoistures,indice,NAA,typ='standard'):
            if indice.get()>0:
                indice.set(indice.get()-1)
                massmanagementcom(P,paginazionespettri,paginazionemasse,paginazioneuncmasse,paginazionemoistures,paginazioneuncmoistures,indice,NAA,typ)
                        
        def massmanagementcom(P,paginazionespettri,paginazionemasse,paginazioneuncmasse,paginazionemoistures,paginazioneuncmoistures,indice,NAA,typ='standard'):#help
            """Display mass management for sample or standard"""
            def update_masses_moisture(MB,MM,ids,valuei,NAA,BT,UMB,UMM):
                try:
                    float(MB.get())
                    NAA.comparatormasses[ids.get()*int(rows)+valuei] = float(MB.get())
                    float(MM.get())
                    NAA.comparatormoisture[ids.get()*int(rows)+valuei] = float(MM.get())
                    float(UMB.get())
                    NAA.comparatoruncertaintymasses[ids.get()*int(rows)+valuei] = float(UMB.get())
                    float(UMM.get())
                    NAA.comparatoruncertaintymoisture[ids.get()*int(rows)+valuei] = float(UMM.get())
                except:
                    pass
                BT.configure(text=f'{round(f_mass(MB,MM),7)} g')
                
            def update_masses_samples_moisture(MB,MM,ids,valuei,NAA,BT,UMB,UMM):
                try:
                    float(MB.get())
                    NAA.samplemasses[ids.get()*int(rows)+valuei] = float(MB.get())
                    float(MM.get())
                    NAA.samplemoisture[ids.get()*int(rows)+valuei] = float(MM.get())
                    float(UMB.get())
                    NAA.sampleuncertaintymasses[ids.get()*int(rows)+valuei] = float(UMB.get())
                    float(UMM.get())
                    NAA.sampleuncertaintymoisture[ids.get()*int(rows)+valuei] = float(UMM.get())
                except:
                    pass
                BT.configure(text=f'{round(f_mass(MB,MM),7)} g')
                
            if indice.get()>-1 and indice.get()<len(paginazionespettri):
                cdn=P.winfo_children()
                for i in cdn:
                    i.destroy()
                F=Frame(P)
                chrw=10
                Bmeno=Button(F, text='<', relief='flat', command=lambda P=P,paginazionespettri=paginazionespettri,paginazionemasse=paginazionemasse,paginazioneuncmasse=paginazioneuncmasse,paginazionemoistures=paginazionemoistures,paginazioneuncmoistures=paginazioneuncmoistures,indice=indice,NAA=NAA,typ=typ : masses_minus(P,paginazionespettri,paginazionemasse,paginazioneuncmasse,paginazionemoistures,paginazioneuncmoistures,indice,NAA,typ)).pack(side=LEFT)
                L=Label(F, text=f'page {indice.get()+1} of {len(paginazionespettri)}').pack(side=LEFT)
                Bpiu=Button(F, text='>', relief='flat', command=lambda P=P,paginazionespettri=paginazionespettri,paginazionemasse=paginazionemasse,paginazioneuncmasse=paginazioneuncmasse,paginazionemoistures=paginazionemoistures,paginazioneuncmoistures=paginazioneuncmoistures,indice=indice,NAA=NAA,typ=typ : masses_plus(P,paginazionespettri,paginazionemasse,paginazioneuncmasse,paginazionemoistures,paginazioneuncmoistures,indice,NAA,typ)).pack(side=LEFT)
                F.pack()
                titles=['filename','mass / g','u(mass) / g','mst. / %','u(mst.) / %']
                FL=Frame(P)
                L=Label(FL, width=1).pack(side=LEFT)
                L=Label(FL, text=titles[0], width=chrw*5, anchor=W).pack(side=LEFT)
                L=Label(FL, text=titles[1], width=chrw).pack(side=LEFT)
                L=Label(FL, text=titles[2], width=chrw).pack(side=LEFT)
                L=Label(FL, width=3).pack(side=LEFT)
                L=Label(FL, text=titles[3], width=chrw-3).pack(side=LEFT)
                L=Label(FL, text=titles[4], width=chrw).pack(side=LEFT)
                FL.pack(anchor=W)
                valuei=0
                for j in range(len(paginazionespettri[indice.get()])):
                    FL=Frame(P)
                    L=Label(FL, width=1).pack(side=LEFT)
                    L=Label(FL, text=paginazionespettri[indice.get()][j].spectrumpath.split('/')[-1], width=chrw*5, anchor=W).pack(side=LEFT) # filename
                    masspinbox=Spinbox(FL, from_=0.001, to=100, width=10, increment=0.001)
                    masspinbox.pack(side=LEFT, padx=2)
                    masspinbox.delete(0,END)
                    masspinbox.insert(END,paginazionemasse[indice.get()][j])
                    uncmasspinbox=Spinbox(FL, from_=0.000001, to=1, width=10, increment=0.0001)
                    uncmasspinbox.pack(side=LEFT, padx=2)
                    uncmasspinbox.delete(0,END)
                    uncmasspinbox.insert(END,paginazioneuncmasse[indice.get()][j])
                    L=Label(FL, width=2).pack(side=LEFT)
                    moisutepinbox=Spinbox(FL, from_=0.01, to=100, width=7, increment=1.0)
                    moisutepinbox.pack(side=LEFT, padx=2)
                    moisutepinbox.delete(0,END)
                    moisutepinbox.insert(END,paginazionemoistures[indice.get()][j])
                    uncmoisutepinbox=Spinbox(FL, from_=0.01, to=100, width=7, increment=0.1)
                    uncmoisutepinbox.pack(side=LEFT, padx=2)
                    uncmoisutepinbox.delete(0,END)
                    uncmoisutepinbox.insert(END,paginazioneuncmoistures[indice.get()][j])
                    L=Label(FL, width=2).pack(side=LEFT)
                    vl = f_mass(masspinbox,moisutepinbox)
                    L_total_mass=Button(FL, text=f'{round(vl,7)} g', width=12)
                    L_total_mass.pack(side=LEFT, padx=5)
                    FL.pack(anchor=W)
                    if typ == 'sample':
                        L_total_mass.configure(command= lambda MB=masspinbox,MM=moisutepinbox,ids=indice,valuei=valuei,NAA=NAA,BT=L_total_mass,UMB=uncmasspinbox,UMM=uncmoisutepinbox : update_masses_samples_moisture(MB,MM,ids,valuei,NAA,BT,UMB,UMM))
                    else:
                        L_total_mass.configure(command= lambda MB=masspinbox,MM=moisutepinbox,ids=indice,valuei=valuei,NAA=NAA,BT=L_total_mass,UMB=uncmasspinbox,UMM=uncmoisutepinbox : update_masses_moisture(MB,MM,ids,valuei,NAA,BT,UMB,UMM))
                    valuei=valuei+1
                L=Label(P).pack()
        
        def menage_sample_masses(NAA):
            """Display single spectrum"""
            if NAA.sample!=None:
                ratiooflines=int(rows)
                residuallenS=len(NAA.sample)
                i=0
                paginazionespettri=[]
                paginazionemasse=[]
                paginazioneuncmasse=[]
                paginazionemoistures=[]
                paginazioneuncmoistures=[]
                while residuallenS>0:
                    try:
                        paginazionespettro=NAA.sample[ratiooflines*i:ratiooflines*i+ratiooflines]
                        paginazionemassa=NAA.samplemasses[ratiooflines*i:ratiooflines*i+ratiooflines]
                        paginazioneuncmassa=NAA.sampleuncertaintymasses[ratiooflines*i:ratiooflines*i+ratiooflines]
                        paginazionemoisture=NAA.samplemoisture[ratiooflines*i:ratiooflines*i+ratiooflines]
                        paginazioneuncmoisture=NAA.sampleuncertaintymoisture[ratiooflines*i:ratiooflines*i+ratiooflines]
                    except IndexError:
                        paginazionespettro=NAA.sample[ratiooflines*i:]
                        paginazionemassa=NAA.samplemasses[ratiooflines*i:]
                        paginazioneuncmassa=NAA.sampleuncertaintymasses[ratiooflines*i:]
                        paginazionemoisture=NAA.samplemoisture[ratiooflines*i:]
                        paginazioneuncmoisture=NAA.sampleuncertaintymoisture[ratiooflines*i:]
                    paginazionespettri.append(paginazionespettro)
                    paginazionemasse.append(paginazionemassa)
                    paginazioneuncmasse.append(paginazioneuncmassa)
                    paginazionemoistures.append(paginazionemoisture)
                    paginazioneuncmoistures.append(paginazioneuncmoisture)
                    residuallenS=residuallenS-ratiooflines
                    i=i+1
                SC=Toplevel()
                SC.title(NAA.sample[0].identity+' mass management')
                SC.resizable(False,False)
                SC.focus()
                L=Label(SC).pack()
                F=Frame(SC)
                P=Frame(F)
                P.pack()
                indice=IntVar(SC)
                indice.set(0)
                massmanagementcom(P,paginazionespettri,paginazionemasse,paginazioneuncmasse,paginazionemoistures,paginazioneuncmoistures,indice,NAA,'sample')
                F.pack()
            
        def manage_masses(NAA,BT):
            if BT.cget('text')=='Standard':
                menage_comparator_masses(NAA)
            elif BT.cget('text')=='Sample':
                menage_sample_masses(NAA)
                    
        def select_one_nuclide(vb,CB,NASN,Lb):
            if vb.get()!='':
                NASN.append(vb.get())
            else:
                NASN.remove(CB.cget('onvalue'))
            Lb.configure(text=f'{len(NASN)} elements')
        
        def save_current_selection(NASN,preset,ENE):
            if len(NASN)!=0 and ENE.get()!='':
                n=0
                nomefile=f'data/presets/{ENE.get()}.sel'
                pool=os.listdir('data/presets')
                while nomefile[13:] in pool:
                    n+=1
                    nomefile=f'data/presets/{ENE.get()}_{n}.sel'
                f=open(nomefile,'w')
                for ks in NASN:
                    f.write(ks)
                    f.write(' ')
                f.close()
                h=os.listdir('data/presets')
                values=[]
                for t in h:
                    if t[-4:]=='.sel':
                        values.append(t[:-4])
                values.append('')
                preset['values']=values
                nomefile=nomefile.replace('data/presets/','')
                nomefile=nomefile.replace('.sel','')
                if nomefile in values:
                    preset.set(nomefile)
                else:
                    preset.set('')
                ENE.delete(0,END)
                ENE.insert(0,'New_saved_selection')
                
        def delete_preset(preset):
            default_list=['All','Medium & Long-lived (50 elements)','Rare earths (16 elements)','Short & Medium-lived (54 elements)','']
            if preset.get() not in default_list:
                os.remove('data/presets/'+preset.get()+'.sel')
                h=os.listdir('data/presets')
                values=[]
                for t in h:
                    if t[-4:]=='.sel':
                        values.append(t[:-4])
                values.append('')
                preset['values']=values
                preset.set('')
            
        def select_preset(ACB,NASN,Lb,preset):
            if preset.get()!='':
                f=open('data/presets/'+preset.get()+'.sel','r')
                r=f.readlines()
                f.close()
                presets=str.split(r[0])
            else:
                presets=[]
            for t in range(len(NASN)):
                NASN.pop(0)
            for txt in presets:
                NASN.append(txt)
            for k in ACB:
                if k.cget('text') in NASN:
                    k.select()
                else:
                    k.deselect()
            Lb.configure(text=f'{len(NASN)} elements')
                    
        def select_nuclides_k0(NASN,Lb):
            lbs=[]
            for i in A:
                lbs.append(i[1])
            lbs=sorted(set(lbs))
            SN=Toplevel()
            SN.title('Element selection')
            SN.resizable(False,False)
            L=Label(SN).pack()
            SN.focus()
            i=0
            ACB=[]
            while i < len(lbs):
                F=Frame(SN)
                L=Label(F, text='', width=1).pack(side=LEFT)
                for k in range(10):
                    try:
                        vb=StringVar(F)
                        CB = Checkbutton(F, text=lbs[i], variable=vb, onvalue=lbs[i], offvalue='', width=5, anchor=W)
                        CB.pack(side=LEFT)
                        if lbs[i] in NASN:
                            CB.select()
                        else:
                            CB.deselect()
                        CB.configure(command=lambda vb=vb,CB=CB,NASN=NASN,Lb=Lb: select_one_nuclide(vb,CB,NASN,Lb))
                        ACB.append(CB)
                        i+=1
                    except IndexError:
                        break
                F.pack(anchor=W)
            L=Label(SN).pack()
            
            F=Frame(SN, pady=2)
            L=Label(F, text='', width=1).pack(side=LEFT)
            L=Label(F, text='filename', width=10, anchor=W).pack(side=LEFT)
            ENW_sel=Entry(F, width=39)
            ENW_sel.pack(side=LEFT)
            ENW_sel.insert(0,'New_saved_selection')
            L=Label(F, text='', width=1).pack(side=LEFT)
            B_save=Button(F, text='Save', width=8, padx=2)
            B_save.pack(side=LEFT)
            F.pack(anchor=W)
            
            F=Frame(SN)
            h=os.listdir('data/presets')
            values=[]
            for t in h:
                if t[-4:]=='.sel':
                    values.append(t[:-4])
            values.append('')
            L=Label(F, text='', width=1).pack(side=LEFT)
            L=Label(F, text='selection', width=10, anchor=W).pack(side=LEFT)
            preset_selectionCB=ttk.Combobox(F, values=values, state='readonly', width=36)
            preset_selectionCB.pack(side=LEFT)
            preset_selectionCB.set('')
            L=Label(F, text='', width=1).pack(side=LEFT)
            B_pre=Button(F, text='Recall', width=8)
            B_pre.pack(side=LEFT)
            B_pre.configure(command=lambda ACB=ACB,NASN=NASN,Lb=Lb,preset=preset_selectionCB : select_preset(ACB,NASN,Lb,preset))
            B_predel=Button(F, text='Delete', width=8)
            B_predel.pack(side=LEFT)
            B_predel.configure(command=lambda preset=preset_selectionCB : delete_preset(preset))
            F.pack(anchor=W)
            B_save.configure(command=lambda NASN=NASN,preset=preset_selectionCB,ENE=ENW_sel : save_current_selection(NASN,preset,ENE))
            L=Label(SN).pack()
            
        def listeffy():
            return [file[:-4] for file in os.listdir('data/efficiencies') if file[-4:].lower() == '.eff']
        
        def savepartialhistory(phist,HS):
            p_copy = phist.copy()
            p_copy['date'] = pd.to_datetime(phist['date'], format='%d-%b-%Y')
            fileTypes = [('Excel file', '.xlsx')]
            nomefile=asksaveasfilename(filetypes=fileTypes, defaultextension='.xlsx')
            if nomefile!=None and nomefile!='':
                p_copy.to_excel(nomefile, index=False)
                HS.destroy()
                
        def selectionecomboselectedplot(event,box,CLB,ax,f,canvas,cov_factor=2):
            p,covp = CLB.recall_calibration(box.get())
            fit = CalibrationFit(p,covp)
            CLB.set_master(fit)
            x = np.linspace(0.05,3.1,300)
            yax,uay = CLB.calibration_master.fit_with_uncertainty(x)
            ll = len(ax.lines)
            if ll != 0:
                for jkl in range(ll):
                    ax.lines.pop(0)
            ax.plot(x*1000, yax, color='k', linewidth=0.75)
            ax.plot(x*1000, yax+uay*cov_factor, color='k', linestyle='--', linewidth=0.5)
            ax.plot(x*1000, yax-uay*cov_factor, color='k', linestyle='--', linewidth=0.5)
            ax.set_ylim(0,np.max(yax+uay*cov_factor)*1.05)
            f.tight_layout()
            canvas.draw()
            
        def calib_openpeaklist(LBPL,CLB,Bt):
            Bt.configure(state='disabled')
            pls = searchrptfilesforcalibration()
            for i in pls:
                CLB.add_spects(i)
                LBPL.insert(END,i[0].split('/')[-1])
            LBPL.focus()
        
        def calib_clearpeaklist(LBPL,CLB,Bt):
            LBPL.delete(0,END)
            CLB.newcalibration_peaklist=[]
            Bt.configure(state='disabled')
            
        def select_Certificates(CLB,GS,LC,Bt):
            def select_CCE(vb,CB,CLB,line):
                if vb.get()!='':
                    CLB.newcalibration_emissions.append(line)
                else:
                    CLB.newcalibration_emissions.remove(line)
            def select_Emissions(vb,CLB,GSs,nm):
                if vb.get()!='':
                    TEM = Toplevel()
                    TEM.title(f'{nm} certificate')
                    TEM.focus()
                    L=Label(TEM, text='certificate date: '+GSs.readable_datetime()).pack(anchor=W, padx=10)
                    L=Label(TEM).pack()
                    Gs_emis = [(str(inf1),inf2,str(inf3),str(inf4),str(inf5),str(inf6),str(inf7),GSs.datetime) for inf1,inf2,inf3,inf4,inf5,inf6,inf7 in zip(GSs.energy,GSs.emitter,GSs.activity,GSs.u_activity,GSs.g_yield,GSs.u_g_yield,GSs.decay_constant)]
                    F=Frame(TEM)
                    L=Label(F, text='', width=5).pack(side=LEFT)
                    L=Label(F, text='Energy / keV', width=12).pack(side=LEFT)
                    L=Label(F, text='Emitter', width=12).pack(side=LEFT)
                    L=Label(F, text='Activity / Bq', width=12).pack(side=LEFT)
                    F.pack(padx=10)
                    for line in Gs_emis:
                        F=Frame(TEM)
                        vcce = StringVar(F)
                        CCE = Checkbutton(F, text='', variable=vcce, onvalue='ON', offvalue='', anchor=W)
                        CCE.pack(side=LEFT, padx=5)
                        if line in CLB.newcalibration_emissions:
                            CCE.select()
                        else:
                            CCE.deselect()
                        CCE.configure(command=lambda vb=vcce,CB=CCE,CLB=CLB,line=line : select_CCE(vb,CB,CLB,line))
                        L=Label(F, text=line[0], width=12).pack(side=LEFT)
                        L=Label(F, text=line[1], width=12).pack(side=LEFT)
                        L=Label(F, text=line[2], width=12).pack(side=LEFT)
                        F.pack()
            
            def on_selection(vb,CB,CLB,LC,GS):
                if vb.get() != '':
                    CLB.newcalibration_certficates.append(vb.get())
                    for i in GS:
                        if i.name == vb.get():
                            CLB.newcalibration_emissions += [(str(inf1),inf2,str(inf3),str(inf4),str(inf5),str(inf6),str(inf7),i.datetime) for inf1,inf2,inf3,inf4,inf5,inf6,inf7 in zip(i.energy,i.emitter,i.activity,i.u_activity,i.g_yield,i.u_g_yield,i.decay_constant)]
                            break
                else:
                    CLB.newcalibration_certficates.remove(CB.cget('onvalue'))
                    for i in GS:
                        if i.name == CB.cget('onvalue'):
                            deleta_emi = [(str(inf1),inf2,str(inf3),str(inf4),str(inf5),str(inf6),str(inf7),i.datetime) for inf1,inf2,inf3,inf4,inf5,inf6,inf7 in zip(i.energy,i.emitter,i.activity,i.u_activity,i.g_yield,i.u_g_yield,i.decay_constant)]
                            break
                    for item in deleta_emi:
                        try:
                            CLB.newcalibration_emissions.remove(item)
                        except ValueError:
                            pass
                if len(CLB.newcalibration_certficates) == 1:
                    sents = f'{CLB.newcalibration_certficates[0]} selected'
                else:
                    sents = f'{len(CLB.newcalibration_certficates)} certificates selected'
                LC.configure(text=sents)

            Bt.configure(state='disabled')
            CTS = Toplevel()
            CTS.title('Certificates')
            CTS.resizable(False,False)
            CTS.focus()
            L=Label(CTS).pack()
            i = 0
            while i < len(GS):
                F=Frame(CTS)
                for k in range(5):
                    try:
                        vb = StringVar(F)
                        CB = Checkbutton(F, text='', variable=vb, onvalue=GS[i].name, offvalue='', anchor=W)
                        CB.pack(side=LEFT, padx=5)
                        if GS[i].name in CLB.newcalibration_certficates:
                            CB.select()
                        else:
                            CB.deselect()
                        CB.configure(command=lambda vb=vb,CB=CB,CLB=CLB,LC=LC,GS=GS : on_selection(vb,CB,CLB,LC,GS))
                        BE = Button(F, text = GS[i].name, relief='flat', anchor=W, command= lambda vb=vb,CLB=CLB,GSs=GS[i],nm=GS[i].name: select_Emissions(vb,CLB,GSs,nm))
                        BE.pack(side=LEFT)
                        separator = ttk.Separator(F, orient="vertical")
                        separator.pack(side=LEFT,fill=Y,expand=1, padx=10)
                        i += 1
                    except IndexError:
                        break
                F.pack()
            L=Label(CTS).pack()
        
        def perform_calibration(Edetector,Efile,CLB,ax,f,canvas,Bt):
            def calculate_fit(x,y,order,tol=0.8):
                satisf = None
                ords = [1,0,-1,-2,-3,-4]
                xx = x / 1000
                W = xx[:,np.newaxis]**ords[:order]
                I = np.identity(W.shape[0])
                popt = np.linalg.inv(W.T@W)@(W.T@y)
                resds = y - popt@W.T
                n,k = y.shape[0], W.shape[1]
                pcov = np.linalg.inv((W.T@np.linalg.inv(np.true_divide(1,n-k)*np.dot(resds,resds)*I))@W)
                perr = np.sqrt(np.diag(pcov))
                if np.max(np.abs(perr/popt)) < tol:
                    satisf = True
                return popt,pcov,resds,satisf

            def calculate_efficiency(line):
                """
                Calculate experimental efficiency
                #line[2] = Activity / Bq
                #line[4] = Gamma yield / 1
                #line[6] = Timezzi / s
                #line[7] = Certificate date / datetime
                #line[8] = Net area / 1
                #line[10] = Start acquisition date / datetime
                #line[11] = Real time / s
                #line[12] = Live time / s
                """
                A,Gy,lb,na,tc,tl = line[2],line[4],line[6],line[8],line[11],line[12]
                lb = np.log(2)/lb
                td = (line[10]-line[7]).days*86400+(line[10]-line[7]).seconds
                return np.log((na*lb*tc)/(tl*(1-np.exp(-lb*tc))*np.exp(-lb*td)*A*Gy))

            def extract_from_peaklist(peak_list,date,real,live):
                return [[float(spect[6]),float(spect[8]),float(spect[9]),date,real,live] for spect in peak_list]
            if Edetector.get()!='' and Efile.get()!='' and len(CLB.newcalibration_emissions) > 6:
                all_peaks = []
                for spect in CLB.newcalibration_peaklist:
                    all_peaks += extract_from_peaklist(spect.peak_list,spect.datetime,spect.real_time,spect.live_time)
                peaks_lines = []
                for line in CLB.newcalibration_emissions:
                    for peaks in all_peaks:
                        if float(line[0])+float(tolerance_energy)>peaks[0] and float(line[0])-float(tolerance_energy)<peaks[0]:
                            line2p = [float(line[0]),line[1],float(line[2]),line[3],float(line[4]),line[5],float(line[6]),line[7]]
                            peaks_lines.append(line2p+peaks[1:])
                            break #pay attention to the break, if multiple emissions only the first encountered is selected
                if len(peaks_lines) > 6:
                    ey = np.array([[line[0],calculate_efficiency(line)] for line in peaks_lines])
                    x,ey = ey[:,0],ey[:,1]
                    order = 6
                    satisfaction = None
                    while not satisfaction:
                        results,covariance,ress,satisfaction = calculate_fit(x,ey,order)
                        order -= 1
                        if order < 4:
                            break
                    CLB.newcalibration_results,CLB.newcalibration_covariance=results,covariance
                    for plt_lines in ax.lines:
                        ax.lines.remove(plt_lines)
                    for t_texts in f.texts:
                        f.texts.remove(t_texts)
                    ax.plot(x,ress, linestyle='', marker='o', markerfacecolor='r', markeredgecolor='k', markersize=3, markeredgewidth=0.5)
                    ylims = np.max(np.abs(ress))+0.005
                    ax.set_ylim(-ylims,ylims)
                    polytext = ''
                    line_polytext = 1
                    for value,unc in zip(CLB.newcalibration_results,np.sqrt(np.diag(CLB.newcalibration_covariance))):
                        if line_polytext == 3:
                            polytext += f'a{line_polytext}: {round(value,6)}({int(abs(unc/value*100))}%)     \n'
                        else:
                            polytext += f'a{line_polytext}: {round(value,6)}({int(abs(unc/value*100))}%)     '
                        line_polytext += 1
                    f.text(0,0.025,polytext)
                    f.tight_layout()
                    f.subplots_adjust(bottom=0.225)
                    canvas.draw()
                    Bt.configure(state='active')
        
        def perform_save(CLB,TS,ED,EN,box):
            def today_date():
                month = ['','jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
                date = datetime.date.today()
                return f'{date.day}-{month[date.month]}-{date.year}'
            if CLB.newcalibration_results is not None and CLB.newcalibration_covariance is not None:
                perr = np.sqrt(np.diag(CLB.newcalibration_covariance))
                pcor = np.identity(6)
                for i in range(len(perr)):
                    for k in range(len(perr)):
                        pcor[i,k] = CLB.newcalibration_covariance[i,k]/(perr[i]*perr[k])
                if len(CLB.newcalibration_results) < 6:
                    addzeros = [0.0,0.0]
                    CLB.newcalibration_results = np.append(CLB.newcalibration_results, addzeros[:6-len(CLB.newcalibration_results)])
                    perr = np.append(perr, addzeros[:6-len(perr)])
                with open(f'data/efficiencies/{EN.get()}.eff','w') as f:
                    for x,u in zip(CLB.newcalibration_results,perr):
                        f.write(f'{x} {u}\n')
                    for line in pcor:
                        f.write(f'{line[0]} {line[1]} {line[2]} {line[3]} {line[4]} {line[5]}\n')
                CLB.calibration_history.loc[len(CLB.calibration_history)] = {'detector':ED.get(), 'date':today_date(), 'process':'calibration', 'result':'new', 'calibration_filename':EN.get()}
                CLB.calibration_history.to_csv('data/efficiencies/history.csv', index=False)
                box['values']=listeffy()
                TS.destroy()
            
        def commandnewcalibration(box,CLB,ED):
            wb = xlrd.open_workbook('data/sources/GS/sources.xlsx')
            GS = []
            for wsn in wb.sheet_names():
                ws = wb.sheet_by_name(wsn)
                dts = pd.read_excel('data/sources/GS/sources.xlsx', sheetname=wsn, sheet_name=wsn, skiprows=1)
                GS.append(GSource(wsn,xlrd.xldate_as_tuple(ws.cell(0,0).value, wb.datemode),dts['Energy'],dts['Nuclide'],dts['Activity'],dts['uActivity'],dts['Gyield'],dts['uGyield'],dts['Hlife']))
            TS = Toplevel()
            TS.title('New calibration')
            TS.resizable(False,False)
            TS.focus()
            SUPERF = Frame(TS)
            FFFDATA = Frame(SUPERF)
            L=Label(FFFDATA).pack()
            F=Frame(FFFDATA)#TS
            L=Label(F, text='detector', width=9, anchor=W, padx=10).pack(side=LEFT)
            edecmame=Entry(F, width=25)
            edecmame.pack(side=LEFT)
            edecmame.delete(0,END)
            edecmame.insert(0,ED.get())
            F.pack(anchor=W)
            F=Frame(FFFDATA)
            L=Label(F, text='filename', width=9, anchor=W, padx=10).pack(side=LEFT)
            emame=Entry(F, width=25)
            emame.pack(side=LEFT)
            F.pack(anchor=W)
            L=Label(FFFDATA).pack()
            L=Label(FFFDATA, text='peaklist').pack(anchor=W, padx=10)
            CLB.newcalibration_peaklist=[]
            SF=Frame(FFFDATA)
            Fbox=Frame(SF)
            LBPL=Listbox(Fbox, width=45)
            LBPL.pack(side=LEFT)
            Fbox.pack(side=LEFT, anchor=W, padx=10)
            Bbox=Frame(SF)
            B_openPL=Button(Bbox, text='Open', width=10)
            B_openPL.pack()
            B_clearPL=Button(Bbox, text='Clear', width=10)
            B_clearPL.pack()
            Bbox.pack(side=RIGHT, anchor=NW, padx=10)
            SF.pack(anchor=W)
            L=Label(FFFDATA).pack()
            CLB.newcalibration_certficates=[]
            CLB.newcalibration_emissions=[]
            F=Frame(FFFDATA)
            Bcert=Button(F, text='Certificates', width=10)
            Bcert.pack(side=LEFT, anchor=W, padx=10)
            L_cert=Label(F, text='0 certificates selected')
            L_cert.pack(side=LEFT)
            F.pack(anchor=W)
            FFFDATA.pack(side=LEFT, anchor=W)
            fpp = Figure(figsize=(5.5, 3.2))
            axp=fpp.add_subplot(111)
            FFFPLOT = Frame(SUPERF)
            Figur=Frame(FFFPLOT)
            Figur.pack(side=LEFT,anchor=CENTER, fill=BOTH, expand=1)
            canvasf = FigureCanvasTkAgg(fpp, master=Figur)
            canvasf.draw()
            canvasf.get_tk_widget().pack(side=TOP, fill=BOTH, expand=1)
            axp.set_ylabel(r'$\Delta_\mathrm{r}(\varepsilon)$ / 1')
            axp.set_xlabel(r'$E$ / keV')
            axp.set_xlim(0,2000)
            axp.grid(linestyle='-.')
            fpp.tight_layout()
            canvasf.draw()
            vspe = ttk.Separator(SUPERF, orient="vertical")
            vspe.pack(side=LEFT,fill=Y,expand=1, padx=10)
            FFFPLOT.pack(side=RIGHT, anchor=NE)
            SUPERF.pack(anchor=W)
            L=Label(TS).pack()
            F=Frame(TS)
            Bokl=Button(F, text='Fit', width=10)
            Bokl.pack(side=LEFT)
            B_savefit=Button(F, text='Save', width=10, state='disabled')
            B_savefit.pack(side=LEFT)
            F.pack(anchor=W, padx=10)
            B_openPL.configure(command=lambda LBPL=LBPL,CLB=CLB,Bt=B_savefit: calib_openpeaklist(LBPL,CLB,Bt))
            B_clearPL.configure(command=lambda LBPL=LBPL,CLB=CLB,Bt=B_savefit: calib_clearpeaklist(LBPL,CLB,Bt))
            Bcert.configure(command=lambda CLB=CLB,GS=GS,LC=L_cert,Bt=B_savefit : select_Certificates(CLB,GS,LC,Bt))
            Bokl.configure(command=lambda Edetector=edecmame,Efile=emame,CLB=CLB,ax=axp,f=fpp,canvas=canvasf,Bt=B_savefit: perform_calibration(Edetector,Efile,CLB,ax,f,canvas,Bt))
            B_savefit.configure(command=lambda CLB=CLB,TS=TS,ED=edecmame,EN=emame,box=box: perform_save(CLB,TS,ED,EN,box))
            L=Label(TS).pack()
            
        def check_calibration(CLB,ED,ax,f,canvas):
            if CLB.calibration_master is not None:
                wb = xlrd.open_workbook('data/sources/GS/sources.xlsx')
                GS = []
                for wsn in wb.sheet_names():
                    ws = wb.sheet_by_name(wsn)
                    dts = pd.read_excel('data/sources/GS/sources.xlsx', sheetname=wsn, sheet_name=wsn, skiprows=1)
                    GS.append(GSource(wsn,xlrd.xldate_as_tuple(ws.cell(0,0).value, wb.datemode),dts['Energy'],dts['Nuclide'],dts['Activity'],dts['uActivity'],dts['Gyield'],dts['uGyield'],dts['Hlife']))
                TS = Toplevel()
                TS.title('Check current calibration')
                TS.resizable(False,False)
                TS.focus()
                SUPERF = Frame(TS)
                FFFDATA = Frame(SUPERF)
                L=Label(FFFDATA).pack()
                F=Frame(FFFDATA)
                L=Label(F, text='detector', width=9, anchor=W, padx=10).pack(side=LEFT)
                edecmame=Entry(F, width=25)
                edecmame.pack(side=LEFT)
                edecmame.delete(0,END)
                edecmame.insert(0,ED.get())
                F.pack(anchor=W)
                L=Label(FFFDATA).pack()
                L=Label(FFFDATA, text='peaklist').pack(anchor=W, padx=10)
                CLB.checkcalibration_peaklist=[]
                SF=Frame(FFFDATA)
                Fbox=Frame(SF)
                LBPL=Listbox(Fbox, width=45)
                LBPL.pack(side=LEFT)
                Fbox.pack(side=LEFT, anchor=W, padx=10)
                Bbox=Frame(SF)
                B_openPL=Button(Bbox, text='Open', width=10)
                B_openPL.pack()
                B_clearPL=Button(Bbox, text='Clear', width=10)
                B_clearPL.pack()
                Bbox.pack(side=RIGHT, anchor=NW, padx=10)
                SF.pack(anchor=W)
                L=Label(FFFDATA).pack()
                CLB.checkcalibration_certficates=[]
                CLB.checkcalibration_emissions=[]
                F=Frame(FFFDATA)
                Bcert=Button(F, text='Certificates', width=10)
                Bcert.pack(side=LEFT, anchor=W, padx=10)
                L_cert=Label(F, text='0 certificates selected')
                L_cert.pack(side=LEFT)
                F.pack(anchor=W)
                FFFDATA.pack(side=LEFT, anchor=W)
                FFFPLOT = Frame(SUPERF)
                check_response = pd.DataFrame(columns=['fit E / keV','test E / keV','isotope','zetha score'])
                #check_response.to_string(col_space=16, index=False, justify='left')
                txtc=Text(FFFPLOT, width=60, heigh=18)
                txtc.insert(END,check_response.to_string(col_space=16, index=False, justify='left'))
                txtc.pack(side=LEFT, anchor=NW, fill=X, expand=1)
                scroller = Scrollbar(FFFPLOT, command=txtc.yview)
                txtc.configure(state='disabled', yscrollcommand = scroller.set)
                scroller.pack(side=RIGHT, anchor=N, fill=Y)
                vspe = ttk.Separator(SUPERF, orient="vertical")
                vspe.pack(side=LEFT,fill=Y,expand=1, padx=10)
                FFFPLOT.pack(side=RIGHT, anchor=NE)
                SUPERF.pack(anchor=W)
                L=Label(TS).pack()
                F=Frame(TS)
                Bokl=Button(F, text='Fit', width=10)
                Bokl.pack(side=LEFT)
                B_savefit=Button(F, text='Save', width=10, state='disabled')
                B_savefit.pack(side=LEFT)
                F.pack(anchor=W, padx=10)
    #            B_openPL.configure(command=lambda LBPL=LBPL,CLB=CLB,Bt=B_savefit: calib_openpeaklist(LBPL,CLB,Bt))
    #            B_clearPL.configure(command=lambda LBPL=LBPL,CLB=CLB,Bt=B_savefit: calib_clearpeaklist(LBPL,CLB,Bt))
    #            Bcert.configure(command=lambda CLB=CLB,GS=GS,LC=L_cert,Bt=B_savefit : select_Certificates(CLB,GS,LC,Bt))
    #            Bokl.configure(command=lambda Edetector=edecmame,Efile=emame,CLB=CLB,ax=axp,f=fpp,canvas=canvasf,Bt=B_savefit: perform_calibration(Edetector,Efile,CLB,ax,f,canvas,Bt))
    #            B_savefit.configure(command=lambda CLB=CLB,TS=TS,ED=edecmame,EN=emame,box=box: perform_save(CLB,TS,ED,EN,box))
                L=Label(TS).pack()
        
        def show_history_coomand(B,hist):
            if B.get() in list(hist['detector']):
                part_hist = hist.loc[hist['detector']==B.get()]
                tit = f'Calibration history for detector {B.get()}'
            else:
                part_hist = hist
                tit = 'Calibration history for all detectors'
            HS = Toplevel()
            HS.title(tit)
            txt = part_hist.to_string(col_space=16, index=False, justify='left')
            HS.resizable(True,False)
            TxtF = Frame(HS)
            text = Text(TxtF, width=100)
            text.pack(side=LEFT, anchor=NW, fill=X, expand=1)
            text.insert(END, txt)
            scroller = Scrollbar(TxtF, command=text.yview)
            text.configure(state='disabled', yscrollcommand = scroller.set)
            scroller.pack(side=RIGHT, anchor=N, fill=Y)
            TxtF.pack(anchor=NW, fill=BOTH, expand=1)
            text.focus()
            bottomline=Frame(HS)
            Bsav=Button(bottomline, text='Export', width=10, command=lambda phist=part_hist, HS=HS : savepartialhistory(phist,HS))
            Bsav.pack(side=LEFT)
            bottomline.pack(anchor=W)
            
        def command_deleteitem(box,CLB,ax,f,canvas):
            if box.get()!='':
                if messagebox.askokcancel('Delete entry', f'Do you want to delete {box.get()} calibration?'):
                    os.remove(f'data/efficiencies/{box.get()}.eff')
                    box['values']=listeffy()
                    box.set('')
                    CLB.calibration_master=None
                    ll = len(ax.lines)
                    if ll != 0:
                        for jkl in range(ll):
                            ax.lines.pop(0)
                    f.tight_layout()
                    canvas.draw()
                    
        def view_model():
            """Shows the measurement model and relevant information"""
            CM=Toplevel()
            CM.title('Measurement model')
            CM.resizable(False,False)
            modimg = PhotoImage(file='data/models/model.png')
            L=Label(CM, image=modimg)
            L.image = modimg
            L.pack()
            CM.focus()
            
        def view_energy_database():
            def parse_text_hl(text,df,cln,tol=0.10):
                def hlf_conversion(value,unit):
                    if unit.lower() == 'y':
                        return value*365.24*86400
                    elif unit.lower() == 'd':
                        return value*86400
                    elif unit.lower() == 'h':
                        return value*3600
                    elif unit.lower() == 'min':
                        return value*60
                    else:
                        return value

                units = ['s','min','h','d','y']
                if text.replace(' ','') != '':
                    unit = text.split()[-1]
                    if unit.lower() in units:
                        text = text.replace(unit,'')
                        text = text.replace(' ','')
                        if '--' in text:
                            try:
                                range_min, range_max = text.split('--')
                                range_min, range_max = float(range_min), float(range_max)
                                range_min, range_max = hlf_conversion(range_min,unit), hlf_conversion(range_max,unit)
                            except ValueError:
                                pass
                            else:
                                if range_max<range_min:
                                    range_min, range_max = range_max, range_min
                                return (df[cln] > range_min) & (df[cln] < range_max)
                        else:
                            try:
                                hlf = float(text)
                            except ValueError:
                                pass
                            else:
                                hlf = hlf_conversion(hlf,unit)
                                return (df[cln] > hlf*(1-tol)) & (df[cln] < hlf*(1+tol))

            def parse_text(text,df,cln,tol=0.5):
                if text.replace(' ','') != '':
                    if '--' in text:
                        try:
                            range_min, range_max = text.split('--')
                            range_min, range_max = float(range_min), float(range_max)
                        except ValueError:
                            pass
                        else:
                            if range_max<range_min:
                                range_min, range_max = range_max, range_min
                            return (df[cln] > range_min) & (df[cln] < range_max)
                    else:
                        text = text.replace(',',' ')
                        try:
                            list_energies = [float(value) for value in text.split()]
                        except ValueError:
                            pass
                        else:
                            list_energies.sort()
                            filt = []
                            for i in list_energies:
                                filt.append((df[cln] < i + tol) & (df[cln] > i - tol))
                            filt_t = filt[0]
                            if len(filt) > 1:
                                for i in filt[1:]:
                                    filt_t = filt_t | i
                            return filt_t
                        
            def df_display(df,txt):
                txt.configure(state='normal')
                txt.delete('0.0',END)
                txt.insert(END, df.to_string(index=False, col_space=10, justify='left'))
                txt.configure(state='disabled')
            
            def output_view(DB,tgt,emit,energy,tmez,check,txt):
                cols = ['target', 'isotope', 'E / keV', 'h_life']
                all_filters = []
                if tgt.get()!='':
                    filt_tgt = DB['target'] == tgt.get()
                    all_filters.append(filt_tgt)
                if emit.get()!='':
                    filt_tgt = DB['isotope'] == emit.get()
                    all_filters.append(filt_tgt)
                filter_energy = parse_text(energy.get(),DB,'E / keV',float(tolerance_energy))        
                if filter_energy is not None:
                    all_filters.append(filter_energy)
                filter_halflife = parse_text_hl(tmez.get(),DB,'t1_2 / s')        
                if filter_halflife is not None:
                    all_filters.append(filter_halflife)
                    
                if len(all_filters) == 0:
                    df_display(DB[cols],txt)
                else:
                    filt_t = all_filters[0]
                    if len(all_filters) > 1:
                        for i in all_filters[1:]:
                            if check.get() == '':
                                filt_t = filt_t & i
                            else:
                                filt_t = filt_t | i
                    df_display(DB[filt_t][cols],txt)

            def human_readable_hl(x):
                if x > 34560000:
                    lb = f'{round(x/(365.24*86400),2)} y'
                elif x > 100000:
                    lb = f'{round(x/86400,2)} d'
                elif x > 10000:
                    lb = f'{round(x/3600,2)} h'
                else:
                    lb = f'{round(x,2)} s'
                return lb
            AA = pd.DataFrame(A, columns=['target','isotope','E / keV','t1_2 / s','note'])
            AA['h_life'] = [human_readable_hl(i) for i in AA['t1_2 / s']]
            TP = Toplevel()
            TP.title('Library')
            TP.resizable(False,False)
            F = Frame(TP)
            L = Label(F, text=f'Current library: {database}', anchor=W).pack()
            F.pack(anchor=W, padx=5, pady=5)
            F =Frame(TP)
            L = Label(F, text='Target', anchor=W, width=8).pack(side=LEFT)
            tg_lst = sorted(set(AA['target']))
            tg_lst.append('')
            CBTG = ttk.Combobox(F, width=7, values=tg_lst, state='readonly')
            CBTG.pack(side=LEFT)
            L = Label(F, text='', anchor=W, width=4).pack(side=LEFT)
            L = Label(F, text='Emitter', anchor=W, width=8).pack(side=LEFT)
            emit_lst = sorted(set(AA['isotope']))
            emit_lst.append('')
            CBEMT = ttk.Combobox(F, width=9, values=emit_lst, state='readonly')
            CBEMT.pack(side=LEFT)
            L = Label(F, text='', anchor=W, width=4).pack(side=LEFT)
            L = Label(F, text='Energy / keV', anchor=W, width=10).pack(side=LEFT)
            E_energy = Entry(F, width=25)
            E_energy.pack(side=LEFT)
            L = Label(F, text='', anchor=W, width=4).pack(side=LEFT)
            L = Label(F, text='t1/2', anchor=W, width=5).pack(side=LEFT)
            E_tmez = Entry(F, width=25)
            E_tmez.pack(side=LEFT)
            F.pack(anchor=W, padx=5, pady=5)
            F =Frame(TP)
            CB_var = StringVar(F)
            CB_var.set('')
            CB = Checkbutton(F, text='Union', variable=CB_var, onvalue='Union', offvalue='')
            CB.pack(side=LEFT, anchor=W)
            F.pack(anchor=W, padx=5, pady=5)
            F =Frame(TP)
            B_conf = Button(F, text='Output', width=8)
            B_conf.pack(side=LEFT)
            F.pack(anchor=W, padx=5, pady=5)
            F =Frame(TP)
            txt = Text(F, height=20, state='disabled')
            txt.pack(side=LEFT, fill=X, expand=1)
            scrbl = Scrollbar(F, command=txt.yview)
            txt['yscrollcommand'] = scrbl.set
            scrbl.pack(side=RIGHT, fill=Y)
            F.pack(anchor=W, fill=X, padx=5, pady=5)
            
            B_conf.configure(command=lambda DB=AA, tgt=CBTG, emit=CBEMT, energy=E_energy, tmez=E_tmez, check=CB_var, txt=txt: output_view(DB,tgt,emit, energy,tmez,check,txt))
            colss = ['target', 'isotope', 'E / keV', 'h_life']
            df_display(AA[['target', 'isotope', 'E / keV', 'h_life']],txt)
            TP.focus()
            
        def view_ceritficates():
            def certificate_selected(CB,txt):
                DF = pd.DataFrame(NAA.certificates.get(CB.get()))
                DF = DF.transpose()
                DF.rename({0: 'ppm', 1: 'u(ppm)', 2: 'xx', 3: 'xs'}, axis='columns', inplace=True)
                DF['urel / %'] = DF['u(ppm)'] / DF['ppm'] * 100
                DF['urel / %'] = [round(x,1) for x in DF['urel / %']]
                cols = ['ppm', 'u(ppm)', 'urel / %']
                txt.configure(state='normal')
                txt.delete('0.0',END)
                txt.insert(END, f'{CB.get()} certificate\nconcentrations with 1σ uncertainties\n\n')
                txt.insert(END,DF.to_string(columns=cols, col_space=10))
                txt.configure(state='disabled')
            
            CRTS = Toplevel()
            CRTS.title('Certificates')
            CRTS.resizable(False,False)
            list_certificates = list(NAA.certificates.keys())
            L = Label(CRTS).pack()
            F = Frame(CRTS)
            L = Label(F, width=1).pack(side=LEFT)
            CCBCRTS = ttk.Combobox(F, width=20, values=list_certificates, state='readonly')
            CCBCRTS.pack(side=LEFT)
            L = Label(F, width=1).pack(side=LEFT)
            L = Label(F, text=f'{len(list_certificates)} certificates loaded', width=25, anchor=W).pack(side=LEFT)
            F.pack(anchor=W)
            F = Frame(CRTS)
            txt = Text(F, width=55, height=30, state='disabled')
            vscrollbar = Scrollbar(F, command=txt.yview)
            txt['yscrollcommand'] = vscrollbar.set
            txt.pack(side=LEFT, fill=X)
            vscrollbar.pack(side=RIGHT, fill=Y)
            F.pack(anchor=W, padx=5, pady=10)
            ewent='<<ComboboxSelected>>'
            CCBCRTS.bind(ewent, lambda event=ewent, CB=CCBCRTS, txt=txt: certificate_selected(CB,txt))
            CRTS.focus()
            if len(list_certificates) > 0:
                CCBCRTS.set(list_certificates[0])
                CCBCRTS.event_generate(ewent)
            
        def openconfermametrologicamodule(BBC):
            """Deprecated function"""
            CLB = Calibration()
            CM=Toplevel()
            CM.title(f'{BBC.cget("text")}')
            CM.resizable(False,False)
            L=Label(CM).pack()
            FL = Frame(CM)
            L=Label(FL, width=1).pack(side=LEFT)
            L=Label(FL, text='detector', width=11).pack(side=LEFT)
            E_detectorname=Entry(FL, width=18)
            E_detectorname.pack(side=LEFT)
            L=Label(FL, width=3).pack(side=LEFT)
            B_show=Button(FL, text='Show history', width=10, command=lambda B=E_detectorname,hist=CLB.calibration_history : show_history_coomand(B,hist))
            B_show.pack(side=LEFT)
            FL.pack(anchor=W)
            L=Label(CM).pack()
            f = Figure(figsize=(8, 4.5))
            ax=f.add_subplot(111)
            MainF = Frame(CM)
            Figur=Frame(MainF)
            Figur.pack(side=LEFT,anchor=CENTER, fill=BOTH, expand=1)
            canvas = FigureCanvasTkAgg(f, master=Figur)
            canvas.draw()
            canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=1)
            toolbar = NavigationToolbar2TkAgg(canvas, Figur)
            toolbar.update()
            ax.set_ylabel(r'$\varepsilon$ / 1')
            ax.set_xlabel(r'$E$ / keV')
            ax.set_xlim(0,2000)
            ax.set_ylim(0,None)
            ax.grid(linestyle='-.')
            f.tight_layout()
            canvas.draw()
            
            BttF = Frame(MainF)
            L=Label(BttF, text='recall calibration').pack(anchor=W, padx=13)
            values=listeffy()
            effy_omboBCP=ttk.Combobox(BttF, values=values, state='readonly', width=25)
            effy_omboBCP.pack(anchor=W, padx=13)
            ewent='<<ComboboxSelected>>'
            L=Label(BttF, text='').pack(anchor=W, padx=13)
            Button_new = Button(BttF, text='New', width=10)
            Button_new.pack(anchor=W, padx=13)
            Button_delete = Button(BttF, text='Delete', width=10)
            Button_delete.pack(anchor=W, padx=13)
            Button_check = Button(BttF, text='Check', width=10)
            Button_check.pack(anchor=W, padx=13)
            BttF.pack(side=RIGHT, anchor=N)
            L=Label(BttF, width=12).pack()
            MainF.pack(anchor=NW, fill=BOTH, expand=1)
            
            effy_omboBCP.bind(ewent, lambda event=ewent,box=effy_omboBCP,CLB=CLB,ax=ax,f=f,canvas=canvas : selectionecomboselectedplot(event,box,CLB,ax,f,canvas))
            Button_new.configure(command = lambda box=effy_omboBCP,CLB=CLB,ED=E_detectorname : commandnewcalibration(box,CLB,ED))
            Button_delete.configure(command = lambda box=effy_omboBCP,CLB=CLB,ax=ax,f=f,canvas=canvas: command_deleteitem(box,CLB,ax,f,canvas))
            Button_check.configure(command = lambda CLB=CLB,ED=E_detectorname,ax=ax,f=f,canvas=canvas : check_calibration(CLB,ED,ax,f,canvas))
            
        def split_strip(item):
            nuc,ene=str.split(item)
            N,Ai=str.split(nuc,'-')
            if Ai[-1]=='m':
                Ai=Ai.replace('m','')
                mg=2.0
            else:
                mg=1.0
            ene=float(ene)
            Ai=float(Ai)
            for t in A:
                if N==t[2] and Ai==t[3] and mg==t[4] and ene==t[5]:
                    mnuclide=t
                    break
            return mnuclide
        
        def ordemeprogresso(Plist,idx=22):
            indexlst=[]
            for y in A:
                if y[1] not in indexlst:
                    indexlst.append(y[1])
            ordered_list=[]
            for lx in indexlst:
                for ix in Plist:
                    if ix[idx]==lx:
                        ordered_list.append(ix)
            return ordered_list
            
        def outcome(message):
            NF=Toplevel()
            NF.title('Message')
            NF.resizable(False,False)
            L=Label(NF).pack()
            L=Label(NF, text=message, width=60).pack()
            L=Label(NF).pack()
            NF.focus()
            event="<FocusOut>"
            NF.bind(event, lambda event=event,G=NF : G.destroy())
            
        def import_masses_fromAquarius():
            def extract_certificate(x):
                try:
                    conv_x = str(int(x))
                except ValueError:
                    conv_x = str(x)
                return conv_x
                
            def set_to_0_otherwise(x,intg=False):
                try:
                    if intg == False:
                        return float(x)
                    else:
                        return int(float(x))
                except:
                    return 0
                
            def apply_certificate(spectrum,cert):
                spectrum.selected_certificate = cert
                spectrum.assign_nuclide=['']*len(spectrum.peak_list)
                try:
                    with open('data/presets/'+cert+'.spl','r') as f:
                        r=f.readlines()
                    for t in range(len(r)):
                        r[t]=r[t].replace('\n','')
                except:
                    print(f'the selected {cert} certificate does not exist in the database')
                else:
                    workingenergylist=[]
                    for i in r:
                        mid=str.split(i)
                        workingenergylist.append(float(mid[-1]))
                    sortedwlist = sorted(workingenergylist)
                    sortedwlist = [sortedwlist[0],sortedwlist[-1]]
                    for tiy in range(len(spectrum.peak_list)):
                        if float(spectrum.peak_list[tiy][6])>sortedwlist[0]-3 and float(spectrum.peak_list[tiy][6])<sortedwlist[-1]+3:
                            for pp in range(len(workingenergylist)):
                                if float(spectrum.peak_list[tiy][6])-float(tolerance_energy)<workingenergylist[pp] and float(spectrum.peak_list[tiy][6])+float(tolerance_energy)>workingenergylist[pp]:
                                    spectrum.assign_nuclide[tiy]=r[pp]
                                    break
                
            if NAA.sample != None or NAA.comparator != None:
                types=[('Excel file','.xls'),('Excel file','.xlsx')]
                filename=askopenfilename(filetypes=types)
                if filename != None and filename != '':
                    ws = xlrd.open_workbook(filename)
                    ws = ws.sheet_by_index(0) #sheet_by_name('DATI') alternatively
                    try:
                        len_sample = len(NAA.sample)
                    except TypeError:
                        len_sample = 0
                    try:
                        len_comparator = len(NAA.comparator)
                    except TypeError:
                        len_comparator = 0
                    for ix in range(len_sample):
                        a, b, c, d = set_to_0_otherwise(ws.cell(12+ix,2).value), set_to_0_otherwise(ws.cell(12+ix,3).value), set_to_0_otherwise(ws.cell(12+ix,4).value), set_to_0_otherwise(ws.cell(12+ix,5).value)
                        NAA.samplemasses[ix] = a/1000
                        NAA.sampleuncertaintymasses[ix] = b/1000
                        NAA.samplemoisture[ix] = c
                        NAA.sampleuncertaintymoisture[ix] = d
                    for ix in range(len_comparator):
                        a, b, c, d, cert = set_to_0_otherwise(ws.cell(12+ix,8).value), set_to_0_otherwise(ws.cell(12+ix,9).value), set_to_0_otherwise(ws.cell(12+ix,10).value), set_to_0_otherwise(ws.cell(12+ix,11).value),extract_certificate(ws.cell(12+ix,12).value)
                        NAA.comparatormasses[ix] = a/1000
                        NAA.comparatoruncertaintymasses[ix] = b/1000
                        NAA.comparatormoisture[ix] = c
                        NAA.comparatoruncertaintymoisture[ix] = d
                        if cert in NAA.certificates.keys():
                            apply_certificate(NAA.comparator[ix],cert)
                    NAA.imported_day, NAA.imported_month, NAA.imported_year, NAA.imported_hours, NAA.imported_minutes, NAA.imported_lenght, NAA.imported_irradiation_code = set_to_0_otherwise(ws.cell(5,2).value,True), set_to_0_otherwise(ws.cell(5,3).value,True), set_to_0_otherwise(ws.cell(5,4).value,True), set_to_0_otherwise(ws.cell(5,6).value,True), set_to_0_otherwise(ws.cell(5,7).value,True), set_to_0_otherwise(ws.cell(7,11).value,True), str(ws.cell(7,3).value)
                    NAA.experimental_additional_information['geometry'] = str(ws.cell(7,5).value)
                    NAA.experimental_additional_information['distance'] = str(ws.cell(7,7).value)
                    NAA.experimental_additional_information['detector'] = str(ws.cell(7,9).value)
                    NAA.imported_channelname = str(ws.cell(8,3).value)
                    if NAA.imported_lenght==0:
                        NAA.imported_lenght = 1
                    try:
                        NAA.imported_irradiation_code = int(float(NAA.imported_irradiation_code))
                    except:
                        pass
                    message_expl = f'Imported information for {len_sample} samples\nand {len_comparator} standards\n& irradiation details\n'
                    messagebox.showinfo('Import complete', message_expl)
            else:
                message_expl = '\nImport spectra before use this command\n'
                messagebox.showwarning('Missing data', message_expl)
        
        def check_allright(EMC,EUMC,EMS,EUMS,EDDC,EDDS,EMUD,EUMUD,EGTHC,EUGTHC,EGEC,EUGEC,EGTHS,EUGTHS,ECOIC,EUCOIC,EWC,EUWC,NAA):
            try:
                float(EMC.get())
                float(EUMC.get())
                float(EMS.get())
                float(EUMS.get())
                float(EDDC.get())
                float(EDDS.get())
                float(EMUD.get())
                float(EUMUD.get())
                float(EGTHC.get())
                float(EUGTHC.get())
                float(EGEC.get())
                float(EUGEC.get())
                float(EGTHS.get())
                float(EUGTHS.get())
                float(ECOIC.get())
                float(EUCOIC.get())
                float(EWC.get())
                float(EUWC.get())
            except:
                print('Invalid data entered\na) values in the spin-boxes should have floating point data type')
            else:
                if float(EMC.get())>0 and float(EMS.get())>0 and float(EWC.get()):
                    NAA.masses=[float(EMC.get()),float(EUMC.get()),float(EMS.get()),float(EUMS.get())]
                    NAA.ddcomparator,NAA.ddsample=float(EDDC.get()),float(EDDS.get())
                    NAA.detector_mu=[float(EMUD.get()),float(EUMUD.get())]
                    NAA.comparatorselfshieldingth=[float(EGTHC.get()),float(EUGTHC.get())]
                    NAA.comparatorselfshieldingepi=[float(EGEC.get()),float(EUGEC.get())]
                    NAA.comparatorCOI=[float(ECOIC.get()),float(EUCOIC.get())]
                    NAA.sampleselfshieldingth=[float(EGTHS.get()),float(EUGTHS.get())]
                    NAA.comparatormassfraction=[float(EWC.get()),float(EUWC.get())]
                    if NAA.comparator!=None and NAA.irradiation!=None and NAA.sample!=None and NAA.standard_comparator!=[None,None] and NAA.efficiencycomparatorfit!=None:#Add something here
                        fileTypes = [('Excel file', '.xlsx')]
                        nomefile=asksaveasfilename(filetypes=fileTypes, defaultextension='.xlsx')
                        if nomefile!=None and nomefile!='':
                            try:
                                do_everything(NAA,nomefile)
                                outcome('Excel file successfully created!')
                                print('Complete!')
                            except:
                                outcome('Some error occurred! Incomplete file saved')
                                NAA.enegysamplefit,NAA.fwhmsamplefit,NAA.efficiencysamplefit,NAA.dersamplefit=None,None,None,[None,None]
                    elif NAA.efficiencycomparatorfit==None:
                        outcome('Detector calibration is required')
                    elif NAA.standard_comparator==[None,None]:
                        outcome('Emission for comparator has not been selected')
                    else:
                        outcome('Check irradiation, sample and standard data selection')
                else:
                    outcome('sample and standard masses, and monitor mass fraction\nmust be greater than 0')
        
        heightsize=10
        logo = PhotoImage(file='k0log.gif')
        F=Frame(M)
        F1=Frame(F)
        L=Label(F1, text='', width=1).pack(side=LEFT)
        L=Label(F1, image=logo)
        L.image = logo
        L.pack(side=LEFT)
        F1.pack(anchor=W, side=LEFT, fill=X, expand=1)
        F2=Frame(F)
        L=Label(F2, text='', width=1).pack(side=RIGHT)
        BHPLMS=Button(F2, text='Input masses', width=13, command=lambda : import_masses_fromAquarius())
        BHPLMS.pack(side=RIGHT)
        #BHPL=Button(F2, text='Input workbook', width=13, command=lambda : create_HyperLabmanually())
        #BHPL.pack(side=RIGHT)
        BSTTGS=Button(F2, text='Settings', width=8, command=lambda M=M : settings_modifications(M))
        BSTTGS.pack(side=RIGHT)
        F2.pack(anchor=E, side=RIGHT, fill=X, expand=1)
        F.pack(pady=heightsize, fill=X, expand=1)
        separator = ttk.Separator(M, orient="horizontal")
        separator.pack(fill=X,expand=1)
        
        #Calibration
        F=Frame(M)
        L=Label(F, text='', width=1).pack(side=LEFT)
        BBC=Button(F, text='Measurement model', width=20)
        BBC.pack(side=LEFT)
        #BBC.configure(command=lambda BBC=BBC : openconfermametrologicamodule(BBC))
        BBC.configure(command=lambda : view_model())
        BCERTIFICATES=Button(F, text='Certificates', width=20)
        BCERTIFICATES.pack(side=LEFT)
        BCERTIFICATES.configure(command=lambda : view_ceritficates())
        BLIBRARY=Button(F, text='Emission library', width=20)
        BLIBRARY.pack(side=LEFT)
        BLIBRARY.configure(command=lambda : view_energy_database())
        F.pack(anchor=W, pady=heightsize)
        separator = ttk.Separator(M, orient="horizontal")
        separator.pack(fill=X,expand=1)
        
        #Background
        F=Frame(M)
        L=Label(F, text='', width=1).pack(side=LEFT)
        BB=Button(F, text='Background', width=10)
        BB.pack(side=LEFT)
        L=Label(F, text='filename:', width=8).pack(side=LEFT)
        LBB=Label(F, width=40, anchor=W)
        LBB.pack(side=LEFT)
        L=Label(F, text='peaks:', width=5).pack(side=LEFT)
        LPKS=Label(F, width=5, anchor=W)
        LPKS.pack(side=LEFT)
        L=Label(F, text='start:', width=6).pack(side=LEFT)
        LDT=Label(F, width=20, anchor=W)
        LDT.pack(side=LEFT)
        BTCB=Button(F, text='Peak list', width=8, command= lambda BB=BB,NAA=NAA : overlook(BB,NAA))
        BTCB.pack(side=LEFT, anchor=W)
        BB.configure(command=lambda BB=BB,NAA=NAA,LBB=LBB,LPKS=LPKS,LDT=LDT : openpeaklistandspectrum(BB,NAA,LBB,LPKS,LDT))
        F.pack(anchor=W, pady=heightsize)
        separator = ttk.Separator(M, orient="horizontal")
        separator.pack(fill=X,expand=1)
        
        #Irradiation
        F=Frame(M)
        L=Label(F, text='', width=1).pack(side=LEFT)
        BIR=Button(F, text='Irradiation', width=10)
        BIR.pack(side=LEFT)
        L=Label(F, text='channel:', width=8).pack(side=LEFT)
        LCH=Label(F, width=22)
        LCH.pack(side=LEFT)
        L=Label(F, text='f / 1:', width=5).pack(side=LEFT)
        LF=Label(F, width=8, anchor=W)
        LF.pack(side=LEFT)
        L=Label(F, text='α / 1:', width=5).pack(side=LEFT)
        LALF=Label(F, width=8, anchor=W)
        LALF.pack(side=LEFT)
        L=Label(F, text='end irr:', width=6).pack(side=LEFT)
        LIDT=Label(F, width=20, anchor=W)
        LIDT.pack(side=LEFT)
        L=Label(F, text='ti / s:', width=5).pack(side=LEFT)
        LITM=Label(F, width=12, anchor=W)
        LITM.pack(side=LEFT)
        F.pack(anchor=W, pady=heightsize)
        BIR.configure(command=lambda BIR=BIR,NAA=NAA,LCH=LCH,LF=LF,LALF=LALF,LIDT=LIDT,LITM=LITM : irradiation_info(BIR,NAA,LCH,LF,LALF,LIDT,LITM))
        
        separator = ttk.Separator(M, orient="horizontal")
        separator.pack(fill=X,expand=1)
        
        #Standard
        FB=Frame(M)
        F=Frame(FB)
        L=Label(F, text='', width=1).pack(side=LEFT)
        BC=Button(F, text='Standard', width=10)
        BC.pack(side=LEFT)
        L=Label(F, text='filename:', width=8).pack(side=LEFT)
        LBC=Label(F, width=40, anchor=W)
        LBC.pack(side=LEFT)
        L=Label(F, text='peaks:', width=5).pack(side=LEFT)
        LPKSC=Label(F, width=5, anchor=W)
        LPKSC.pack(side=LEFT)
        L=Label(F, text='start:', width=6).pack(side=LEFT)
        LDTC=Label(F, width=20, anchor=W)
        LDTC.pack(side=LEFT)
        BTCC=Button(F, text='Peak list', width=8, command= lambda BC=BC,NAA=NAA : overlook(BC,NAA))
        BTCC.pack(side=LEFT)
        BTCSAll=Button(F, text='Clear', width=8, command= lambda NAA=NAA,LBC=LBC,LPKSC=LPKSC,LDTC=LDTC,BT=BC : clearall(NAA,LBC,LPKSC,LDTC,BT))
        BTCSAll.pack(side=LEFT)
        BC.configure(command=lambda BC=BC,NAA=NAA,LBC=LBC,LPKSC=LPKSC,LDTC=LDTC : openpeaklistandspectrum(BC,NAA,LBC,LPKSC,LDTC))
        F.pack(anchor=W)
        F=Frame(FB)
        L=Label(F, text='', width=12).pack(side=LEFT)
        MCMC=Button(F, text='Masses', width=10, command = lambda NAA=NAA,BT=BC : manage_masses(NAA,BT))
        MCMC.pack(side=LEFT)
        F.pack(anchor=W)
        
        FB.pack(anchor=W, pady=heightsize)

        separator = ttk.Separator(M, orient="horizontal")
        separator.pack(fill=X,expand=1)
        
        #Sample
        FB=Frame(M)
        F=Frame(FB)
        L=Label(F, text='', width=1).pack(side=LEFT)
        BS=Button(F, text='Sample', width=10)
        BS.pack(side=LEFT)
        L=Label(F, text='filename:', width=8).pack(side=LEFT)
        LBS=Label(F, width=40, anchor=W)
        LBS.pack(side=LEFT)
        L=Label(F, text='peaks:', width=5).pack(side=LEFT)
        LPKSS=Label(F, width=5, anchor=W)
        LPKSS.pack(side=LEFT)
        L=Label(F, text='start:', width=6).pack(side=LEFT)
        LDTS=Label(F, width=20, anchor=W)
        LDTS.pack(side=LEFT)
        BTCS=Button(F, text='Peak list', width=8, command= lambda BS=BS,NAA=NAA : overlook(BS,NAA))
        BTCS.pack(side=LEFT)
        BTCAll=Button(F, text='Clear', width=8, command= lambda NAA=NAA,LBS=LBS,LPKSS=LPKSS,LDTS=LDTS,BT=BS : clearall(NAA,LBS,LPKSS,LDTS,BT))
        BTCAll.pack(side=LEFT)
        BS.configure(command=lambda BS=BS,NAA=NAA,LBS=LBS,LPKSS=LPKSS,LDTS=LDTS : openpeaklistandspectrum(BS,NAA,LBS,LPKSS,LDTS))
        L=Label(F, text='', width=1).pack(side=LEFT)
        F.pack(anchor=W)
        F=Frame(FB)
        L=Label(F, text='', width=12).pack(side=LEFT)
        MCMS=Button(F, text='Masses', width=10, command=lambda NAA=NAA,BT=BS : manage_masses(NAA,BT))
        MCMS.pack(side=LEFT)
        F.pack(anchor=W)
        
        FB.pack(anchor=W, pady=heightsize)
        
        separator = ttk.Separator(M, orient="horizontal")
        separator.pack(fill=X,expand=1)
        
        #Elaborate
        FB=Frame(M)
        F=Frame(FB)
        L=Label(F, text='', width=1).pack(side=LEFT)
        Bopti=Button(F, text='Elaborate', width=10)
        F.pack(anchor=W)
        FB.pack(anchor=W, pady=heightsize)
        
        Bopti.pack(side=LEFT)
        Bopti.configure(command=lambda NAA=NAA : menage_comparator_emissions(NAA))
    
    def on_closing():
        if messagebox.askokcancel('Quit R-LENA', 'Unsaved data will be lost.\n\nDo you want to quit?'):
            M.destroy()
    M=Tk()
    M.title('Main')
    M.resizable(False,False)
    M.protocol("WM_DELETE_WINDOW", on_closing)
    mainscreen(M,NAA)
    M.mainloop()

if __name__=='__main__':
    main()
