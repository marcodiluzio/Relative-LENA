# -*- coding: utf-8 -*-
"""
Created on Fri Apr  3 17:55:38 2020

@author: Marco Di Luzio
"""

from tkinter import *
import os
from tkinter import ttk
from tkinter import messagebox
import numpy as np
import pandas as pd

class DataManager:
    def __init__(self,M,version=0.2,date='30 September 2020'):
        self.VERSION = version
        self.VERSION_DATE = date
        self.folder_CRM = 'data/sources'
        self.folder_library = 'data/rel_data'
        #self.folder_gammasources = 'data/sources/GS/sources.xlsx'
        self.certificates = {}
        
        self.all_elements = ['Ag', 'Al', 'Ar', 'As', 'Au', 'Ba', 'Bi', 'Br', 'Ca', 'Cd', 'Ce', 'Cl', 'Co', 'Cr', 'Cs', 'Cu', 'Dy', 'Er', 'Eu', 'F', 'Fe', 'Ga', 'Gd', 'Ge', 'Hf', 'Hg', 'Ho', 'I', 'In', 'Ir', 'K', 'La', 'Lu', 'Mg', 'Mn', 'Mo', 'Na', 'Nb', 'Nd', 'Ni', 'Os', 'Pb', 'Pd', 'Pr', 'Pt', 'Rb', 'Re', 'Rh', 'Ru', 'S', 'Sb', 'Sc', 'Se', 'Si', 'Sm', 'Sn', 'Sr', 'Ta', 'Tb', 'Te', 'Th', 'Ti', 'Tl', 'Tm', 'U', 'V', 'W', 'Xe', 'Y', 'Yb', 'Zn', 'Zr']
        self.all_cols = ['CnV','SD','n','Notes']
        
        self._splashscreen()
        
    def _welcome(self):
        print(f'Database Manager for R-LENA\nversion {self.VERSION} ({self.VERSION_DATE})')
        
    def _back_to_splashscreen(self):
        cdn=M.winfo_children()
        for i in cdn:
            i.destroy()
        self._splashscreen(False)
        
    def _get_database_emiss(self,filename='data/kimp0-01r.txl'):
        with open(filename,'r') as f:
            r = f.readlines()
        return r[0].replace('\n','')
        
    def _splashscreen(self,welcome=True):
        if welcome == True:
            self._welcome()
        F = Frame(M)
        L = Label(F, text='What are you eager to manage today?', anchor=W).pack(padx=5, pady=5)
        B_CRM = Button(F, text='Certificates', width=20, command=lambda : self._module_CRM())
        B_CRM.pack(padx=5, pady=2)
        B_Emlibrary = Button(F, text='Emission library', width=20, command=lambda : self._module_Emlibrary(self._get_database_emiss()))
        B_Emlibrary.pack(padx=5, pady=2)
        #B_Gsources = Button(F, text='Gamma sources', width=10)
        #B_Gsources.pack(padx=5)
        F.pack()
        
    def _module_CRM(self):
        cdn=M.winfo_children()
        for i in cdn:
            i.destroy()
        listset = [file[:-5] for file in os.listdir(self.folder_CRM) if file[-5:].lower()=='.xlsx']
        F = Frame(M)
        L = Button(F, text='←', width=3, command=lambda : self._back_to_splashscreen()).pack(anchor=W, padx=5, pady=5)
        L = Label(F, text='Select an existing certificate database to modify from the list or create a new one', anchor=W).pack(padx=5)
        Fs = Frame(F)
        self.combo_CRM = ttk.Combobox(Fs, width=20, values=listset)
        self.combo_CRM.pack(side=LEFT)
        L = Label(Fs, text='', width=5).pack(side=LEFT)
        B_create = Button(Fs, text='Create', width=10, command=lambda : self.create_new())
        B_create.pack(side=LEFT)
        B_delete = Button(Fs, text='Delete', width=10, command=lambda : self.delete_database())
        B_delete.pack(side=LEFT)
        B_modify = Button(Fs, text='Open', width=10, command=lambda : self.open_modify())
        B_modify.pack(side=LEFT)
        B_savecur = Button(Fs, text='Save', width=10, command=lambda : self.save_database())
        B_savecur.pack(side=LEFT)
        Fs.pack(anchor=W, padx=5)
        F_CRM = Frame(F, borderwidth=0.75, relief=SUNKEN)
        
        L = Label(F_CRM, text='Certificate', anchor=W).pack(anchor=W)
        F_line = Frame(F_CRM)
        self.certificate_combo = ttk.Combobox(F_line, width=20, state='readonly', values=[])
        self.certificate_combo.pack(side=LEFT)
        L = Label(F_line, text='', width=5).pack(side=LEFT)
        B_add_cert = Button(F_line, text='Add', width=10)
        B_add_cert.pack(side=LEFT)
        B_delete_cert = Button(F_line, text='Delete', width=10)
        B_delete_cert.pack(side=LEFT)
        B_rename = Button(F_line, text='Rename', width=10)
        B_rename.pack(side=LEFT)
        F_line.pack(anchor=W)
        
        F_line = Frame(F_CRM)
        L = Label(F_line, text='Element', width=15, anchor=W).pack(side=LEFT)
        L = Label(F_line, text='x / ppm', width=15, anchor=W).pack(side=LEFT)
        L = Label(F_line, text='u(x) / ppm', width=15, anchor=W).pack(side=LEFT)
        L = Label(F_line, text='n', width=15, anchor=W).pack(side=LEFT)
        L = Label(F_line, text='Notes', width=15, anchor=W).pack(side=LEFT)
        F_line.pack(anchor=W)
        F_line = Frame(F_CRM)
        self.element_combo = ttk.Combobox(F_line, width=12, state='readonly', values=[])
        self.element_combo.pack(side=LEFT)
        self.x_spinbox = Spinbox(F_line, width=10, from_=0, to=1000000, increment=1)
        self.x_spinbox.pack(side=LEFT, padx=15)
        self.ux_spinbox = Spinbox(F_line, width=10, from_=0, to=1000000, increment=1)
        self.ux_spinbox.pack(side=LEFT, padx=15)
        self.n_spinbox = Spinbox(F_line, width=10, from_=0, to=1000000, increment=1)
        self.n_spinbox.pack(side=LEFT, padx=15)
        self.E_notes = Entry(F_line, width=20)
        self.E_notes.pack(side=LEFT, padx=10)
        F_line.pack(anchor=W)

        F_CRM.pack(padx=5, pady=15)
        F.pack()
        self.loading_progress = ttk.Progressbar(M, orient = HORIZONTAL, length = 100, mode = 'determinate') 
        
        B_add_cert.configure(command=lambda : self.add_certificate())
        B_delete_cert.configure(command=lambda : self.delete_certificate())
        B_rename.configure(command=lambda : self.rename_certificate())
        
        event = '<FocusIn>'
        events = '<<ComboboxSelected>>'
        self.combo_CRM.bind(event,lambda event : self.avoid_quirks())
        self.certificate_combo.bind(events,lambda event : self.avoid_quirks_2())
        self.element_combo.bind(events,lambda events : self.show_selection())
        
        eventent = '<Return>'
        eventtab = '<Tab>'
        self.x_spinbox.bind(eventent,lambda event=eventent,SBX=self.x_spinbox,idx=0 : self._modify_certificate(SBX,idx))
        self.x_spinbox.bind(eventtab,lambda event=eventtab,SBX=self.x_spinbox,idx=0 : self._modify_certificate(SBX,idx))
        self.ux_spinbox.bind(eventent,lambda event=eventent,SBX=self.ux_spinbox,idx=1 : self._modify_certificate(SBX,idx))
        self.ux_spinbox.bind(eventtab,lambda event=eventtab,SBX=self.ux_spinbox,idx=1 : self._modify_certificate(SBX,idx))
        self.n_spinbox.bind(eventent,lambda event=eventent,SBX=self.n_spinbox,idx=2 : self._modify_certificate(SBX,idx))
        self.n_spinbox.bind(eventtab,lambda event=eventtab,SBX=self.n_spinbox,idx=2 : self._modify_certificate(SBX,idx))
        self.E_notes.bind(eventent,lambda event=eventent,SBX=self.E_notes,idx=3 : self._modify_certificate(SBX,idx))
        self.E_notes.bind(eventtab,lambda event=eventtab,SBX=self.E_notes,idx=3,BS=self.x_spinbox : self._modify_certificate(SBX,idx,BS))
        
    def avoid_quirks_2(self):
        self.element_combo.set(self.element_combo['values'][0])
        self.show_selection()
        
    def avoid_quirks(self):
        self.certificate_combo['values'] = []
        self.certificate_combo.set('')
        self.element_combo['values'] = []
        self.element_combo.set('')
        
    def _decrypts(self,value,types='num'):
        if types=='str':
            value = str(value)
            if value.lower() == 'nan':
                return ''
            else:
                return value
        else:
            if np.isnan(value):
                return '0'
            else:
                return value
            
    def delete_database(self):
        if self.combo_CRM.get()!='' and self.combo_CRM.get() in self.combo_CRM['values']:
            with open('data/kimp0-01r.txl','r') as f:
                settings_selection = f.readlines()[1]
            settings_selection = settings_selection.replace('\n','')
            if f'{self.combo_CRM.get()}.xlsx' == settings_selection:
                messagebox.showwarning('Warning','you cannot delete the database currently used in R-LENA!\nChange the selected database from R-LENA settings before deleting it')
            else:
                if len(self.combo_CRM['values']) > 1:
                    if messagebox.askokcancel('Delete',f'Are you sure you want to delete\nthis certificate database ({self.combo_CRM.get()})?'):
                        os.remove(f'data/sources/{self.combo_CRM.get()}.xlsx')
                        self.combo_CRM['values'] = [file[:-5] for file in os.listdir(self.folder_CRM) if file[-5:]=='.xlsx']
                        self.combo_CRM.set('')
                        self.certificate_combo['values'] = []
                        self.certificate_combo.set('')
                        self.element_combo['values'] = []
                        self.element_combo.set('')
                        self._clean_presets()
                elif len(self.combo_CRM['values']) == 1:
                    messagebox.showwarning('Warning','You need at least one database in your collection!')
    
    def save_database(self):
        if self.certificate_combo.get()!='':
            self.loading_progress.pack(anchor=W, padx=5)
            self.loading_progress['maximum'] = len(self.certificates.keys())
            self.loading_progress['value'] = 0
            with pd.ExcelWriter(f'data/sources/{self.combo_CRM.get()}.xlsx') as writer:
                for key, value in self.certificates.items():
                    value.to_excel(writer, sheet_name=key, index_label='Elements')
                    self.loading_progress['value'] += 1
                    self.loading_progress.update()
            self.loading_progress.forget()
            self._clean_presets()
            messagebox.showinfo('Info',f'Modification to {self.combo_CRM.get()} database\nwere successfully saved')
        else:
            messagebox.showinfo('Info','You can only save modifications performed\non the currently open database\n\nSelect a database and open it to start modifying')
                        
    def add_certificate(self):
        if self.certificate_combo.get()!='':
            nn = 0
            while f'new_certificate_{nn}' in self.certificates.keys():
                nn += 1
            self.certificates[f'new_certificate_{nn}'] = pd.DataFrame(index=self.all_elements, columns=self.all_cols)
            self.certificate_combo['values'] = [key for key in self.certificates.keys()]
            self.certificate_combo.set(f'new_certificate_{nn}')
            self.element_combo.set('')
            
    def delete_certificate(self):
        if self.certificate_combo.get()!='':
            if len(self.certificates) > 1:
                if messagebox.askokcancel('Delete',f'Are you sure you want to delete\nthis certificate ({self.certificate_combo.get()})?'):
                    self.certificates.pop(self.certificate_combo.get())
                    self.certificate_combo['values'] = [key for key in self.certificates.keys()]
                    self.certificate_combo.set(self.certificate_combo['values'][0])
                    self.element_combo.set('')
                    self._clean_presets()
            elif len(self.certificates) == 1:
                messagebox.showwarning('Warning','A database needs at least 1 certificate.\nDelete the whole database instead')
    
    def rename_certificate(self):
        if self.certificate_combo.get()!='':
            h,w,a,b=self.certificate_combo.winfo_height(),self.certificate_combo.winfo_width(),self.certificate_combo.winfo_rootx(),self.certificate_combo.winfo_rooty()
            TLRN=Toplevel()
            TLRN.geometry(f'{w}x{h}+{a}+{b+h}')
            TLRN.overrideredirect(True)
            TLRN.resizable(False,False)
            E=Entry(TLRN)
            E.pack(side=LEFT, fill=X, expand=True)
            E.insert(0,self.certificate_combo.get())
            E.focus()            
            BC=Button(TLRN, text='ok', command=lambda E=E,TLRN=TLRN : self._confirm_rename(E,TLRN)).pack(side=RIGHT)
            ew='<Return>'
            E.bind(ew,lambda ew=ew,E=E,TLRN=TLRN : self._return_confirm_rename(E,TLRN))
            event="<FocusOut>"
            TLRN.bind(event,lambda event=event,MW=TLRN : MW.destroy())
            
    def _confirm_rename(self,E,TLRN):
        if E.get()!='' and E.get()!=self.certificate_combo.get() and E.get() not in self.certificates.keys():
            value = self.certificates.pop(self.certificate_combo.get())
            self.certificates[E.get()] = value
            self.certificate_combo['values'] = [key for key in self.certificates.keys()]
            self.certificate_combo.set(E.get())
            self._clean_presets()
        else:
            if E.get()=='' or E.get() in self.certificates.keys():
                messagebox.showerror('Error','Rename was unsuccessful.\n\nCauses might be:\n- an invalid new name was entered\n- a certificate with some name than new name was already present in the current database')
        TLRN.destroy()
    
    def _return_confirm_rename(self,E,TLRN):
        self._confirm_rename(E,TLRN)
            
    def show_selection(self):
        if self.element_combo.get()!='':
            value = self._decrypts( self.certificates[self.certificate_combo.get()].loc[self.element_combo.get(),self.certificates[self.certificate_combo.get()].columns[0]])
            unc = self._decrypts( self.certificates[self.certificate_combo.get()].loc[self.element_combo.get(),self.certificates[self.certificate_combo.get()].columns[1]])
            ns = self._decrypts( self.certificates[self.certificate_combo.get()].loc[self.element_combo.get(),self.certificates[self.certificate_combo.get()].columns[2]])
            note = self._decrypts( self.certificates[self.certificate_combo.get()].loc[self.element_combo.get(),self.certificates[self.certificate_combo.get()].columns[3]],'str')
            self.x_spinbox.delete(0,END)
            self.x_spinbox.insert(END,value)
            self.ux_spinbox.delete(0,END)
            self.ux_spinbox.insert(END,unc)
            self.n_spinbox.delete(0,END)
            self.n_spinbox.insert(END,ns)
            self.E_notes.delete(0,END)
            self.E_notes.insert(END,note)
            
    def _modify_certificate(self,SBX,idx,BS=None):
        if self.certificate_combo.get()!= '' and self.element_combo.get()!='':
            newvalue = self._get_new_value(SBX,idx)
            if newvalue is not None:
                self.certificates[self.certificate_combo.get()].loc[self.element_combo.get(),self.certificates[self.certificate_combo.get()].columns[idx]] = newvalue
        if BS is not None:
            self.slide_focus(BS)
            return "break"
        
    def slide_focus(self,BS):
        idx = self.element_combo['values'].index(self.element_combo.get())
        if idx < len(self.element_combo['values'])-1:
            self.element_combo.set(self.element_combo['values'][idx+1])
            self.element_combo.event_generate('<<ComboboxSelected>>')
            BS.focus()
        else:
            self.element_combo.focus()
            
    def _get_new_value(self,SBX,idx):
        if idx!=3:
            try:
                value = float(SBX.get())
            except:
                return None
            else:
                if value!=self.certificates[self.certificate_combo.get()].loc[self.element_combo.get(),self.certificates[self.certificate_combo.get()].columns[idx]] and value > 0 and value < 1.000000001e6:
                    return value
        else:
            if SBX.get()!=self.certificates[self.certificate_combo.get()].loc[self.element_combo.get(),self.certificates[self.certificate_combo.get()].columns[idx]]:
                    return SBX.get()
        
    def create_new(self):
        if self.combo_CRM.get()!='':
            if self.combo_CRM.get() in self.combo_CRM['values']:
                messagebox.showerror('Error','A certificate database with same name already exists')
            else:
                self.new_database()
                self.certificate_combo['values'] = []
                self.certificate_combo.set('')
                self.element_combo['values'] = []
                self.element_combo.set('')
                self._clean_presets()
                messagebox.showinfo('Info',f'New certificate database ({self.combo_CRM.get()})\nwas successfully created!')
        else:
            messagebox.showinfo('Info','Insert the name for the new database')
    
    def open_modify(self):
        if self.combo_CRM.get()!='' and self.combo_CRM.get() in self.combo_CRM['values']:
            self.certificates, elems = self.open_databases(self.combo_CRM.get())
            self.certificate_combo['values'] = [key for key in self.certificates.keys()]
            self.certificate_combo.set(self.certificate_combo['values'][0])
            self.element_combo['values'] = elems
            self.element_combo.set('')
        else:
            messagebox.showinfo('Info','Select a database to modify from the left combobox')
            
    def new_database(self):
        new_d = {}
        new_d['new_certificate_0'] = pd.DataFrame(index=self.all_elements, columns=self.all_cols)
        with pd.ExcelWriter(f'data/sources/{self.combo_CRM.get()}.xlsx') as writer:
            for key, value in new_d.items():
                value.to_excel(writer, sheet_name=key, index_label='Elements')
        
        self.combo_CRM['values'] = [file[:-5] for file in os.listdir(self.folder_CRM) if file[-5:]=='.xlsx']
        
    def open_databases(self,filename):
        import xlrd
        
        wb = xlrd.open_workbook(f'data/sources/{filename}.xlsx')
        certificates = {}
        column_line = []
        self.loading_progress.pack(anchor=W, padx=5)
        self.loading_progress['maximum'] = len(wb.sheet_names())
        self.loading_progress['value'] = 0
        for sname in wb.sheet_names():
            fs = wb.sheet_by_name(sname)
            values = pd.read_excel(wb,sheet_name=sname, index_col=0)
            values.astype({'Notes': 'str'})
            values.index = [item.strip() for item in values.index]
            column_line += fs.col_values(0,1,fs.nrows)
            self.loading_progress['value'] += 1
            self.loading_progress.update()
            certificates[sname] = values
        column_line = sorted(set([item.strip() for item in column_line]))
        self.loading_progress.forget()
        return certificates, column_line
    
    def _clean_presets(self):
        for filename in [filesel for filesel in os.listdir('data/presets/') if filesel[-4:]=='.spl']:
            os.remove('data/presets/'+filename)
    
    def _module_Emlibrary(self,relative_database):
        cdn=M.winfo_children()
        for i in cdn:
            i.destroy()
        self.energy_df = pd.read_excel(f'{self.folder_library}/{relative_database}', sheet_name='analytical')
        listset = [f'{emitter} {energy}' for emitter, energy in zip(self.energy_df[self.energy_df.columns[1]],self.energy_df[self.energy_df.columns[2]])]
        F = Frame(M)
        L = Button(F, text='←', width=3, command=lambda : self._back_to_splashscreen()).pack(anchor=W, padx=5, pady=5)
        L = Label(F, text=f'Emission database: {self._get_database_emiss()}', anchor=W).pack(padx=5)
        L = Label(F, text='Select an existing emission to delete or create a new one', anchor=W).pack(padx=5, pady=5)
        Fs = Frame(F)
        self.combo_ENE = ttk.Combobox(Fs, width=20, values=listset, state='readonly')
        self.combo_ENE.pack(side=LEFT)
        L = Label(Fs, text='', width=5).pack(side=LEFT)
        BE_create = Button(Fs, text='Add', width=10)
        BE_create.configure(command=lambda BT=BE_create : self.create_new_entry(BT))
        BE_create.pack(side=LEFT)
        BE_delete = Button(Fs, text='Delete', width=10, command=lambda : self.delete_entry())
        BE_delete.pack(side=LEFT)
        BE_savecur = Button(Fs, text='Save', width=10, command=lambda : self.save_Edatabase())
        BE_savecur.pack(side=LEFT)
        Fs.pack(anchor=W, padx=5)
        FF = Frame(F)
        FF.pack(pady=10)
        VSF = Frame(F)
        L = Label(VSF, text='target', width=10, anchor=W).pack(side=LEFT)
        L = Label(VSF, text='emitter', width=10, anchor=W).pack(side=LEFT)
        L = Label(VSF, text='energy / keV', width=10, anchor=W).pack(side=LEFT)
        L = Label(VSF, text='half-life', width=10, anchor=W).pack(side=LEFT)
        VSF.pack(anchor=W, padx=5, pady=2)
        VSF = Frame(F)
        self.Ltarget = Label(VSF, text='', width=10, anchor=W)
        self.Ltarget.pack(side=LEFT)
        self.Lemitter = Label(VSF, text='', width=10, anchor=W)
        self.Lemitter.pack(side=LEFT)
        self.Lenergy = Label(VSF, text='', width=10, anchor=W)
        self.Lenergy.pack(side=LEFT)
        self.LHL = Label(VSF, text='', width=10, anchor=W)
        self.LHL.pack(side=LEFT)
        VSF.pack(anchor=W, padx=5, pady=2)
        FF = Frame(F)
        FF.pack(pady=10)
        F.pack()
        
        self.loading_progress = ttk.Progressbar(M, orient = HORIZONTAL, length = 100, mode = 'determinate')
        
        event = '<<ComboboxSelected>>'
        self.combo_ENE.bind(event,lambda event : self.cb_selection())
        
    def create_new_entry(self,BT):
        h,w,a,b=BT.winfo_height(),BT.winfo_width(),BT.winfo_rootx(),BT.winfo_rooty()
        TLRN=Toplevel()
        TLRN.geometry(f'{w*4}x{h*6}+{a}+{b+h}')
        TLRN.overrideredirect(True)
        TLRN.resizable(False,False)
        F = Frame(TLRN)
        Fline = Frame(F)
        l = Label(Fline, text='target', width=15, anchor=W).pack(side=LEFT)
        self.cb_target = ttk.Combobox(Fline, width=10, values=self.all_elements)
        self.cb_target.pack(side=RIGHT)
        Fline.pack(fill=X, anchor=W)
        Fline = Frame(F)
        l = Label(Fline, text='emitter', width=15, anchor=W).pack(side=LEFT)
        self.cb_emitter_A = ttk.Combobox(Fline, width=10, values=[str(i) for i in range(300)])
        self.cb_emitter_A.pack(side=RIGHT)
        self.cb_emitter = ttk.Combobox(Fline, width=10, values=self.all_elements)
        self.cb_emitter.pack(side=RIGHT)
        Fline.pack(fill=X, anchor=W)
        Fline = Frame(F)
        l = Label(Fline, text='energy / kev', width=15, anchor=W).pack(side=LEFT)
        self.cb_energy = Spinbox(Fline, from_=0.00, to=4000, increment=0.01, width=10)
        self.cb_energy.pack(side=RIGHT)
        Fline.pack(fill=X, anchor=W)
        Fline = Frame(F)
        l = Label(Fline, text='half-life', width=15, anchor=W).pack(side=LEFT)
        self.cb_unit = ttk.Combobox(Fline, width=4, values=['y','d','m','h','s'], state='readonly')
        self.cb_unit.pack(side=RIGHT)
        self.cb_unit.set('d')
        self.cb_halflife = Spinbox(Fline, from_=0.00, to=10000, increment=0.01, width=10)
        self.cb_halflife.pack(side=RIGHT)
        Fline.pack(fill=X, anchor=W)
        Fline = Frame(F)
        B_accept = Button(Fline, text='Confirm', width=10, command=lambda MW=TLRN : self.accept_values(MW))
        B_accept.pack(side=LEFT)
        B_cancel = Button(Fline, text='Cancel', width=10, command=lambda MW=TLRN : MW.destroy())
        B_cancel.pack(side=RIGHT)
        Fline.pack(pady=6)
        F.pack(anchor=W, padx=10, pady=10, fill=BOTH)
        TLRN.focus()
    
    def accept_values(self,MW):
        try:
            HL, EN = float(self.cb_halflife.get()), float(self.cb_energy.get())
        except:
            HL, EN = 0, 0
        if self.cb_target.get() in self.cb_target['values'] and self.cb_emitter.get() in self.cb_emitter['values'] and self.cb_emitter_A.get() in self.cb_emitter_A['values'] and HL > 0 and EN > 0:
            emittr = f'{self.cb_emitter.get()}-{self.cb_emitter_A.get()}'
            HL = self._to_second(HL,self.cb_unit.get())
            update_df = pd.DataFrame([[self.cb_target.get(),emittr,EN,HL]], columns=['Target','Emitter','Energy / keV','Half-life / s'])
            self.energy_df = self.energy_df.append(update_df, ignore_index=True)
            self.energy_df.sort_values('Energy / keV', ascending=False, inplace=True, kind='quicksort', na_position='last', ignore_index=True)
            MW.destroy()
            listset = [f'{emitter} {energy}' for emitter, energy in zip(self.energy_df[self.energy_df.columns[1]],self.energy_df[self.energy_df.columns[2]])]
            self.combo_ENE['values'] = listset
            self.combo_ENE.set(self.combo_ENE['values'][0])
            self.cb_selection()
        
    def _to_second(self,HL,unit):
        if unit=='y':
            return HL*86400*365.24
        elif unit=='d':
            return HL*86400
        elif unit=='h':
            return HL*3600
        elif unit=='m':
            return HL*60
        else:
            return HL
    
    def delete_entry(self):
        if self.combo_ENE.get()!='' and len(self.combo_ENE['values']) > 1:
            if messagebox.askokcancel('Delete',f'Are you sure to delete the emission {self.combo_ENE.get()}?'):
                emit, energy = self.combo_ENE.get().split()
                filt = (self.energy_df[self.energy_df.columns[1]] == emit) & (self.energy_df[self.energy_df.columns[2]] == float(energy))
                lidx = self.energy_df[filt].index[0]
                self.energy_df.drop(labels=lidx, inplace=True)
                self.combo_ENE['values'] = [f'{emitter} {energy}' for emitter, energy in zip(self.energy_df[self.energy_df.columns[1]],self.energy_df[self.energy_df.columns[2]])]
                self.combo_ENE.set('')
                self.Ltarget.configure(text='')
                self.Lemitter.configure(text='')
                self.Lenergy.configure(text='')
                self.LHL.configure(text='')
        else:
            if self.combo_ENE.get()=='':
                messagebox.showwarning('Warning','Select emission to delete')
            elif len(self.combo_ENE['values']) == 1:
                messagebox.showwarning('Warning','At least 1 emission is needed')
    
    def save_Edatabase(self):
        self.energy_df.to_excel(f'{self.folder_library}/relative_database.xlsx', sheet_name='analytical', index=False)
        self._clean_presets()
        messagebox.showinfo('Info','Emission database saved!')
    
    def cb_selection(self):
        if self.combo_ENE.get()!='':
            emit, energy = self.combo_ENE.get().split()
            filt = (self.energy_df[self.energy_df.columns[1]] == emit) & (self.energy_df[self.energy_df.columns[2]] == float(energy))
            lidx = self.energy_df[filt].index[0]
            self.Ltarget.configure(text=self.energy_df.loc[lidx,self.energy_df.columns[0]])
            self.Lemitter.configure(text=emit)
            self.Lenergy.configure(text=energy)
            self.LHL.configure(text=self.repr_halflife(self.energy_df.loc[lidx,self.energy_df.columns[3]]))
    
    def repr_halflife(self,value):
        if value > 32000000:
            return f'{round(value/(86400*365.24),3)} y'
        elif value > 100000:
            return f'{round(value/(86400),2)} d'
        elif value > 3600:
            return f'{round(value/(3600),2)} h'
        elif value > 60:
            return f'{round(value/(60),2)} m'
        else:
            return f'{round(value,2)} s'

VERSION = 0.2
VERSION_DATE = '8 October 2020'

if __name__ == '__main__':
    M = Tk()
    M.title('R-LENA Database Manager')
    M.resizable(False,False)
    DataManager(M,VERSION,VERSION_DATE)
    M.mainloop()
