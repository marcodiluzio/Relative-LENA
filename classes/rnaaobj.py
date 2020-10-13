# -*- coding: utf-8 -*-
"""
Created on Fri Jun  8 10:30:09 2018

@author: Marco Di Luzio
"""

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askopenfilenames
import datetime
import numpy as np
import csv
import os
import pandas as pd
try:
    import xlsxwriter
except ModuleNotFoundError:
    input('Module xlsxwriter not found, install the Anaconda and click the "Add Anaconda to my PATH environment variable" checkbox\nPress any key to exit')
    quit()

class Spectrum:
    """define a spectrum with all attached information."""
    def __init__(self,identity='Test',start_acquisition=datetime.datetime.today(),real_time=1000,live_time=999,peak_list=None,counts=None,path=None):
        self.identity=identity#Identity(Background,Comparator,Analytes) -> str
        self.datetime=start_acquisition#StartAcquisition -> datetime
        self.real_time=real_time#Real time -> float
        self.live_time=live_time#Live time -> float
        self.peak_list=peak_list#HyperLab_peaklist ->list
        self.counts=counts#spectrum ->list
        self.spectrumpath=path
        self.assign_nuclide=None
        
    def deadtime(self,out='str'):
        try:
            deadtime=(self.real_time-self.live_time)/self.real_time
            if out=='str':
                deadtime=deadtime*100
                deadtime=str(deadtime.__round__(2))+' %'
        except:
            if out=='str':
                deadtime='Invalid'
            else:
                deadtime=None
        return deadtime
    
    def readable_datetime(self):
        return self.datetime.strftime("%d/%m/%Y %H:%M:%S")
    
    def number_of_channels(self):
        try:
            len(self.counts)
        except:
            return 0
        else:
            return len(self.counts)
        
    def defined_spectrum_integral(self,start,width):
        if start>0 and start<len(self.counts) and start+width<len(self.counts):
            integr=0
            for i in range(width):
                integr += self.counts[start+i]
            return integr
        else:
            return None
        
    def define(self):
        return self.identity
    
    def filename(self):
        filename=str(self.spectrumpath)
        filename=str.split(filename,'/')
        return filename[-1]
        
class Irradiation:
    """Define irradiation conditions and channel with respective parameters"""
    def __init__(self,irradiation_end,irradiation_time,u_irradiation_time,f,uf=0,alfa=0,ualfa=0,channel_name='Not defined',code='Not defined'):
        self.code=code
        self.channel=channel_name
        self.datetime=irradiation_end
        self.time=int(irradiation_time)
        self.utime=u_irradiation_time
        self.f=f
        self.uf=uf
        self.a=alfa
        self.ua=ualfa
        
    def __repr__(self):
        return 'Irradiation '+str(self.code)+' in channel '+str(self.channel)
    
    def readable_datetime(self):
        return self.datetime.strftime("%d/%m/%Y %H:%M:%S")
        
class Linear_fit:
    """Define linear (E/FWHM) fit"""
    def __init__(self,m,q):
        self.m,self.q=m,q
        
    def fun(self,Ch):
        return self.m*Ch+self.q
    
    def fun_rev(self,E):
        return (E-self.q)/self.m
    
    def fun_squared(self,Ch):
        return np.sqrt(self.m*Ch+self.q)
    
    def __eq__(self,C):
        try:
            if self.m==C.m and self.q==C.q:
                return True
            else:
                return False
        except:
            return False
    
class Poly_fit:
    """Define polynomial efficiency fit"""
    def __init__(self,A1=1,uA1=0,A2=1,uA2=0,A3=1,uA3=0,A4=1,uA4=0,A5=1,uA5=0,A6=1,uA6=0,A11=1,A12=0,A13=0,A14=0,A15=0,A16=0,A21=0,A22=1,A23=0,A24=0,A25=0,A26=0,A31=0,A32=0,A33=1,A34=0,A35=0,A36=0,A41=0,A42=0,A43=0,A44=1,A45=0,A46=0,A51=0,A52=0,A53=0,A54=0,A55=1,A56=0,A61=0,A62=0,A63=0,A64=0,A65=0,A66=1,name=''):
        self.P=np.array([[A1,uA1],[A2,uA2],[A3,uA3],[A4,uA4],[A5,uA5],[A6,uA6]])
        self.C=np.array([[A11,A12,A13,A14,A15,A16],[A21,A22,A23,A24,A25,A26],[A31,A32,A33,A34,A35,A36],[A41,A42,A43,A44,A45,A46],[A51,A52,A53,A54,A55,A56],[A61,A62,A63,A64,A65,A66]])
        self.name=name
        
    def fun(self,E):
        E=E/1000 #E input in keV is converted to MeV
        return np.exp(self.P[0,0]*E+self.P[1,0]+self.P[2,0]*E**(-1)+self.P[3,0]*E**(-2)+self.P[4,0]*E**(-3)+self.P[5,0]*E**(-4))
    
    def __eq__(self,C):
        try:
            if self.P[0,0]==C.P[0,0] and self.P[1,0]==C.P[1,0] and self.P[2,0]==C.P[2,0] and self.P[3,0]==C.P[3,0] and self.P[4,0]==C.P[4,0] and self.P[5,0]==C.P[5,0] and self.P[0,1]==C.P[0,1] and self.P[1,1]==C.P[1,1] and self.P[2,1]==C.P[2,1] and self.P[3,1]==C.P[3,1] and self.P[4,1]==C.P[4,1] and self.P[5,1]==C.P[5,1]:
                return True
            else:
                return False
        except:
            return False
        
class GSource:
    def __init__(self,name,dt,E,Tg,Bq,uBq,gY,ugY,l):
        self.name=name
        self.datetime=datetime.datetime(*dt)
        self.energy=E
        self.emitter=Tg
        self.activity=Bq
        self.u_activity=uBq
        self.g_yield=gY
        self.u_g_yield=ugY
        self.decay_constant=l
        
    def readable_datetime(self):
        date,time=str.split(str(self.datetime))
        year,month,day=str.split(date,'-')
        hour,minute,second=str.split(time,':')
        return (day+'/'+month+'/'+year+' '+hour+':'+minute+':'+second)
    
class Calibration:
    """Define the metrological confirmation of detector calibration"""
    def __init__(self):
        self.calibration_master=None
        self.calibration_check=None
        self.calibration_history=self.find_history()
        self.newcalibration_peaklist=[]
        self.newcalibration_certficates=[]
        self.newcalibration_emissions=[]
        self.newcalibration_results=None
        self.newcalibration_covariance=None
        self.checkcalibration_peaklist=[]
        self.checkcalibration_certficates=[]
        self.checkcalibration_emissions=[]
#        self.checkcalibration_results=None
#        self.checkcalibration_covariance=None
        
    def find_history(self):
        hst = pd.read_csv('data/efficiencies/history.csv')
        return hst
    
    def add_spects(self,spec_info):
        spectrum = Spectrum(identity='Calibration',start_acquisition=spec_info[1],real_time=spec_info[2],live_time=spec_info[3],peak_list=spec_info[4],path=spec_info[0])
        self.newcalibration_peaklist.append(spectrum)
    
    def recall_calibration(self,name):
        with open(f'data/efficiencies/{name}.eff') as f:
            rf = f.readlines()
            p, up = np.loadtxt(rf[:6], unpack=True)
            corp = np.loadtxt(rf[6:])
        covp = np.identity(len(corp))
        for i in range(len(up)):
            for k in range(len(up)):
                covp[i,k] = corp[i,k] * up[i] * up[k]
        return p,covp
    
    def set_master(self,fit):
        self.calibration_master=fit
        
#    def zetha_score(self,energy,x,ux):
#        return (x-mfit)/np.sqrt(np.power(ux,2)+np.power(sfit,2))
        
class CalibrationFit:
    """Define a polynomial calibration fit"""
    def __init__(self,params,cov_matrix):
        self.params = params
        self.cov_matrix = cov_matrix
        
    def get_uncertainty(self):
        return np.sqrt(np.diag(self.cov_matrix))
        
    def fit_with_uncertainty(self,x):
        """x array or single value"""
        def regular_e_function(a,E):
            return np.exp(a[0]*E + a[1] + a[2]*np.power(E,-1) + a[3]*np.power(E,-2) + a[4]*np.power(E,-3) + a[5]*np.power(E,-4))
        up = self.get_uncertainty()
        c_unc = []
        plus = np.copy(self.params)
        minus = np.copy(self.params)
        for e in x:
            ci = np.array([0.0 for n in range(6)])
            for c in range(len(ci)):
                if self.params[c]!=0.0:
                    plus[c] = self.params[c] + up[c]
                    minus[c] = self.params[c] - up[c]
                    ci[c] = (regular_e_function(plus,e)-regular_e_function(minus,e))/(2*up[c])
                    plus[c],minus[c] = self.params[c],self.params[c]
            c_unc.append(np.sqrt(ci@self.cov_matrix@ci.T))
        W=x[:, np.newaxis]**[1,0,-1,-2,-3,-4]
        return np.exp(self.params@W.T),np.array(c_unc)
    
class RNAAnalysis:
    """Define the actual analysis"""
    def __init__(self):
        self.irradiation=None
        self.comparator=None
        self.comparatormasses=None
        self.comparatoruncertaintymasses=None
        self.comparatormoisture=None
        self.comparatoruncertaintymoisture=None
        self.comparatorcertificates=None
        #self.enegycomparatorfit=None
        #self.fwhmcomparatorfit=None
        #self.efficiencycomparatorfit=None
        #self.dercomparatorfit=[None,None]
        #self.ddcomparator=None
        self.sample=None
        self.samplemasses=None
        self.sampleuncertaintymasses=None
        self.samplemoisture=None
        self.sampleuncertaintymoisture=None
        #self.enegysamplefit=None
        #self.fwhmsamplefit=None
        #self.efficiencysamplefit=None
        #self.dersamplefit=[None,None]
        #self.ddsample=None
        #self.masses=[None,None,None,None]
        self.quantification=None
        self.detection_limits_FWHM=3
        self.selected_nuclides=[]
        self.background=None
        #self.enegybackgroundfit=None
        #self.fwhmbackgroundfit=None
        #self.efficiencybackgroundfit=None
        self.standard_comparator=None
        self.relative_method=None
        self.detector_mu=[None,None]
        self.comparatorselfshieldingth=[None,None]
        self.comparatorselfshieldingepi=[None,None]
        self.sampleselfshieldingth=[None,None]
        self.comparatorCOI=[None,None]
        self.comparatormassfraction=[None,None]
        self.default_utdm=None#seconds
        self.default_udeltatd=None#seconds
        self.default_utc=None#seconds
        self.default_uE=None#keV
        self.certificates=None
        self.elem_dataframe=None
        self.experimental_additional_information = {}
        self.info = {}
    
    def set_backgroungspectrum(self,S):
        if type(S)==Spectrum:
            self.background=S
            
    def set_comparatorspectrum(self,S):
        if type(S)==Spectrum:
            try:
                self.comparator.append(S)#lists!
                self.comparatorcertificates.append('')
                self.comparatormasses.append(0.0)
                self.comparatoruncertaintymasses.append(0.0)
                self.comparatormoisture.append(0.0)
                self.comparatoruncertaintymoisture.append(0.0)
            except:
                self.comparator=[]
                self.comparator.append(S)
                self.comparatorcertificates=[]
                self.comparatorcertificates.append('')
                self.comparatormasses=[]
                self.comparatormasses.append(0.0)
                self.comparatoruncertaintymasses=[]
                self.comparatoruncertaintymasses.append(0.0)
                self.comparatormoisture=[]
                self.comparatormoisture.append(0.0)
                self.comparatoruncertaintymoisture=[]
                self.comparatoruncertaintymoisture.append(0.0)
            
    def set_samplespectrum(self,S):
        if type(S)==Spectrum:
            try:
                self.sample.append(S)#lists!
                self.samplemasses.append(0.0)
                self.sampleuncertaintymasses.append(0.0)
                self.samplemoisture.append(0.0)
                self.sampleuncertaintymoisture.append(0.0)
            except:
                self.sample=[]
                self.sample.append(S)
                self.samplemasses=[]
                self.samplemasses.append(0.0)
                self.sampleuncertaintymasses=[]
                self.sampleuncertaintymasses.append(0.0)
                self.samplemoisture=[]
                self.samplemoisture.append(0.0)
                self.sampleuncertaintymoisture=[]
                self.sampleuncertaintymoisture.append(0.0)
                
    def set_irradiation(self,I):
        if type(I)==Irradiation:
            self.irradiation=I
            
    def set_matrix_typeI(self,iline,monitor,nn):
        HDGS=['ti','np,a','λ,a','td,a','tc,a','tl,a','COI,a','w,a','k0,Au(a)','Gth,a','Ge,a','Q0,a','Er,a','np,c','λ,c','td,c','tc,c','tl,c','COI,c','w,c','m,c','k0,Au(c)','Gth,c','Ge,c','Q0,c','Er,c','f','α','A1','A2','A3','A4','A5','A6','Ea','Ec','dεra','dεrc','dda','ddc','µ']
        MP=np.zeros((len(HDGS),2))
        MP[0,0],MP[0,1],MP[1,0],MP[1,1]=float(self.irradiation.time),float(self.irradiation.utime),float(iline[8]),float(iline[9])
        if iline[53]=='M':
            U=60
        elif iline[53]=='H':
            U=3600
        elif iline[53]=='D':
            U=86400
        elif iline[53]=='Y':
            U=86400*365.24
        else:
            U=1
        L=np.log(2)/(float(iline[52])*U)
        uL=float(iline[54])/float(iline[52])*L
        MP[2,0],MP[2,1]=L,uL
        dc=self.sample[nn].datetime-self.comparator.datetime
        td=dc.days*86400+dc.seconds
        MP[3,0],MP[3,1],MP[4,0],MP[4,1],MP[5,0],MP[5,1],MP[6,0],MP[6,1]=td,self.default_udeltatd,self.sample[nn].real_time,self.default_utc,self.sample[nn].live_time,self.default_utc,1.0,0.0
        MP[7,0],MP[7,1]=self.masses[2],self.masses[3]
        if iline[29]!='':
            uk0=iline[29]*iline[28]/100
        else:
            uk0=0.02*iline[28]
        MP[8,0],MP[8,1],MP[9,0],MP[9,1],MP[10,0],MP[10,1]=iline[28],uk0,self.sampleselfshieldingth[0],self.sampleselfshieldingth[1],1.0,0.0
        if iline[97]!='':
            uQ=iline[96]*iline[97]/100
        else:
            uQ=iline[96]*0.2
        MP[11,0],MP[11,1]=iline[96],uQ
        if iline[99]!='':
            uE=iline[98]*iline[99]/100
        else:
            uE=iline[98]*0.5
        MP[12,0],MP[12,1]=iline[98],uE
        MP[13,0],MP[13,1]=float(monitor[8]),float(monitor[9])
        if monitor[53]=='M':
            U=60
        elif monitor[53]=='H':
            U=3600
        elif monitor[53]=='D':
            U=86400
        elif monitor[53]=='Y':
            U=86400*365.24
        else:
            U=1
        L=np.log(2)/(float(monitor[52])*U)
        uL=float(monitor[54])/float(monitor[52])*L
        MP[14,0],MP[14,1]=L,uL
        dc=self.comparator.datetime-self.irradiation.datetime
        td=dc.days*86400+dc.seconds
        MP[15,0],MP[15,1],MP[16,0],MP[16,1],MP[17,0],MP[17,1]=td,self.default_utdm,self.comparator.real_time,0.01,self.comparator.live_time,0.01
        MP[18,0],MP[18,1],MP[19,0],MP[19,1]=self.comparatorCOI[0],self.comparatorCOI[1],self.masses[0],self.masses[1]
        if monitor[29]!='':
            uk0=monitor[29]*monitor[28]/100
        else:
            uk0=0.02*monitor[28]
        MP[20,0],MP[20,1],MP[21,0],MP[21,1],MP[22,0],MP[22,1],MP[23,0],MP[23,1]=self.comparatormassfraction[0],self.comparatormassfraction[1],monitor[28],uk0,self.comparatorselfshieldingth[0],self.comparatorselfshieldingth[1],self.comparatorselfshieldingepi[0],self.comparatorselfshieldingepi[1]
        if monitor[97]!='':
            uQ=monitor[96]*monitor[97]/100
        else:
            uQ=monitor[96]*0.2
        MP[24,0],MP[24,1]=monitor[96],uQ
        if monitor[99]!='':
            uE=monitor[98]*monitor[99]/100
        else:
            uE=monitor[98]*0.5
        MP[25,0],MP[25,1],MP[26,0],MP[26,1],MP[27,0],MP[27,1]=monitor[98],uE,self.irradiation.f,self.irradiation.uf,self.irradiation.a,self.irradiation.ua
        MP[28:34,0:2]=self.efficiencycomparatorfit.P
        MP[34,0],MP[34,1]=iline[26]/1000,self.default_uE/1000
        MP[35,0],MP[35,1]=monitor[26]/1000,self.default_uE/1000
        if self.dercomparatorfit[0] is not None:
            MP[36,0],MP[36,1]=0,0
            for ssq in range(len(self.dercomparatorfit[0])):
                if self.dercomparatorfit[0][ssq]==int(iline[26]):
                    MP[36,0],MP[36,1]=self.dercomparatorfit[1][ssq],0
                    break
        else:
            MP[36,0],MP[36,1]=0,0
        if self.dercomparatorfit[0] is not None:
            MP[37,0],MP[37,1]=0,0
            for ssq in range(len(self.dercomparatorfit[0])):
                if self.dercomparatorfit[0][ssq]==int(monitor[26]):
                    MP[37,0],MP[37,1]=self.dercomparatorfit[1][ssq],0
                    break
        else:
            MP[37,0],MP[37,1]=0,0
        MP[38,0],MP[38,1]=0,self.ddsample
        MP[39,0],MP[39,1]=0,self.ddcomparator
        MP[40,0],MP[40,1]=self.detector_mu[0],self.detector_mu[1]
        #Correlations
        MC=np.identity((len(HDGS)))
        MC[28:28+len(self.efficiencysamplefit.C),28:28+len(self.efficiencysamplefit.C)]=self.efficiencysamplefit.C
        if monitor[98]==iline[98] and monitor[96]==iline[96]:#condition for same emitting isotope
            MC[HDGS.index('λ,a'),HDGS.index('λ,c')],MC[HDGS.index('λ,c'),HDGS.index('λ,a')]=1.0,1.0
            MC[HDGS.index('Q0,a'),HDGS.index('Q0,c')],MC[HDGS.index('Q0,c'),HDGS.index('Q0,a')]=1.0,1.0
            MC[HDGS.index('Er,a'),HDGS.index('Er,c')],MC[HDGS.index('Er,c'),HDGS.index('Er,a')]=1.0,1.0
            if monitor[28]==iline[28] and monitor[26]==iline[26] and monitor[22]==iline[22]:#relative analysis conditions
                MC[HDGS.index('k0,Au(a)'),HDGS.index('k0,Au(c)')],MC[HDGS.index('k0,Au(c)'),HDGS.index('k0,Au(a)')]=1.0,1.0
                MC[HDGS.index('Ea'),HDGS.index('Ec')],MC[HDGS.index('Ec'),HDGS.index('Ea')]=1.0,1.0
                MC[HDGS.index('COI,a'),HDGS.index('COI,c')],MC[HDGS.index('COI,c'),HDGS.index('COI,a')]=1.0,1.0
        return MP,MC
    
    def set_matrix_detectiontypeI(self,iline,monitor,nn):
        HDGS=['ti','np,a','λ,a','td,a','tc,a','tl,a','COI,a','w,a','k0,Au(a)','Gth,a','Ge,a','Q0,a','Er,a','np,c','λ,c','td,c','tc,c','tl,c','COI,c','w,c','m,c','k0,Au(c)','Gth,c','Ge,c','Q0,c','Er,c','f','α','A1','A2','A3','A4','A5','A6','Ea','Ec','dεra','dεrc','dda','ddc','µ']
        detectionlimitrange=int(self.detection_limits_FWHM*self.fwhmsamplefit.fun_squared(self.enegysamplefit.fun_rev(iline[5])))+1
        AD=self.sample[nn].defined_spectrum_integral(int(self.enegysamplefit.fun_rev(iline[5])-detectionlimitrange/2),detectionlimitrange)
        if AD!=None and AD>0:
            MP=np.zeros((len(HDGS),2))
            MP[0,0],MP[0,1],MP[1,0],MP[1,1]=float(self.irradiation.time),0.0,2.71+3.29*np.sqrt(AD),0.0
            if iline[32]=='M':
                U=60
            elif iline[32]=='H':
                U=3600
            elif iline[32]=='D':
                U=86400
            elif iline[32]=='Y':
                U=86400*365.24
            else:
                U=1
            L=np.log(2)/(float(iline[31])*U)
            MP[2,0],MP[2,1]=L,0.0
            dc=self.sample[nn].datetime-self.comparator.datetime
            td=dc.days*86400+dc.seconds
            MP[3,0],MP[3,1],MP[4,0],MP[4,1],MP[5,0],MP[5,1],MP[6,0],MP[6,1]=td,0.0,self.sample[nn].real_time,0.0,self.sample[nn].live_time,0.0,1.0,0.0
            MP[7,0],MP[7,1]=self.masses[2],0.0
            MP[8,0],MP[8,1],MP[9,0],MP[9,1],MP[10,0],MP[10,1]=iline[7],0.0,self.sampleselfshieldingth[0],0.0,1.0,0.0
            MP[11,0],MP[11,1]=iline[75],0.0
            MP[12,0],MP[12,1]=iline[77],0.0
            MP[13,0],MP[13,1]=float(monitor[8]),0.0
            if monitor[53]=='M':
                U=60
            elif monitor[53]=='H':
                U=3600
            elif monitor[53]=='D':
                U=86400
            elif monitor[53]=='Y':
                U=86400*365.24
            else:
                U=1
            L=np.log(2)/(float(monitor[52])*U)
            MP[14,0],MP[14,1]=L,0.0
            dc=self.comparator.datetime-self.irradiation.datetime
            td=dc.days*86400+dc.seconds
            MP[15,0],MP[15,1],MP[16,0],MP[16,1],MP[17,0],MP[17,1]=td,0.0,self.comparator.real_time,0.0,self.comparator.live_time,0.0
            MP[18,0],MP[18,1],MP[19,0],MP[19,1]=self.comparatorCOI[0],0.0,self.masses[0],0.0
            MP[20,0],MP[20,1],MP[21,0],MP[21,1],MP[22,0],MP[22,1],MP[23,0],MP[23,1]=self.comparatormassfraction[0],0.0,monitor[28],0.0,self.comparatorselfshieldingth[0],0.0,self.comparatorselfshieldingepi[0],0.0
            MP[24,0],MP[24,1]=monitor[96],0.0
            MP[25,0],MP[25,1],MP[26,0],MP[26,1],MP[27,0],MP[27,1]=monitor[98],0.0,self.irradiation.f,0.0,self.irradiation.a,0.0
            MP[28:34,0:1]=self.efficiencycomparatorfit.P[:,0:1]
            MP[34,0],MP[34,1]=iline[5]/1000,0.0
            MP[35,0],MP[35,1]=monitor[26]/1000,0.0  
            if self.dercomparatorfit[0] is not None:
                MP[36,0],MP[36,1]=0,0
                for ssq in range(len(self.dercomparatorfit[0])):
                    if self.dercomparatorfit[0][ssq]==int(iline[5]):
                        MP[36,0],MP[36,1]=self.dercomparatorfit[1][ssq],0
                        break
            else:
                MP[36,0],MP[36,1]=0,0
            if self.dercomparatorfit[0] is not None:
                MP[37,0],MP[37,1]=0,0
                for ssq in range(len(self.dercomparatorfit[0])):
                    if self.dercomparatorfit[0][ssq]==int(monitor[26]):
                        MP[37,0],MP[37,1]=self.dercomparatorfit[1][ssq],0
                        break
            else:
                MP[37,0],MP[37,1]=0,0
            MP[38,0],MP[38,1]=0,0.0
            MP[39,0],MP[39,1]=0,0.0
            MP[40,0],MP[40,1]=self.detector_mu[0],0.0      
            #Correlations
            MC=np.identity((len(HDGS)))
            MC[28:28+len(self.efficiencysamplefit.C),28:28+len(self.efficiencysamplefit.C)]=self.efficiencysamplefit.C
            if monitor[98]==iline[77] and monitor[96]==iline[75]:#condition for same emitting isotope
                MC[HDGS.index('λ,a'),HDGS.index('λ,c')],MC[HDGS.index('λ,c'),HDGS.index('λ,a')]=1.0,1.0
                MC[HDGS.index('Q0,a'),HDGS.index('Q0,c')],MC[HDGS.index('Q0,c'),HDGS.index('Q0,a')]=1.0,1.0
                MC[HDGS.index('Er,a'),HDGS.index('Er,c')],MC[HDGS.index('Er,c'),HDGS.index('Er,a')]=1.0,1.0
                if monitor[28]==iline[7] and monitor[26]==iline[5] and monitor[22]==iline[1]:#condition for relative analysis
                    MC[HDGS.index('k0,Au(a)'),HDGS.index('k0,Au(c)')],MC[HDGS.index('k0,Au(c)'),HDGS.index('k0,Au(a)')]=1.0,1.0
                    MC[HDGS.index('Ea'),HDGS.index('Ec')],MC[HDGS.index('Ec'),HDGS.index('Ea')]=1.0,1.0
                    MC[HDGS.index('COI,a'),HDGS.index('COI,c')],MC[HDGS.index('COI,c'),HDGS.index('COI,a')]=1.0,1.0
        else:
            MP,MC=None,None
        return MP,MC
            
    def define_matrix(self,iline,monitor,nn):
        if iline[43]=='I' or iline[43]=='IVB' or iline[43]=='IIB' or iline[43]=='VI':
            MP,MC=self.set_matrix_typeI(iline,monitor,nn)
        else:
            MP,MC=None,None
        return MP,MC
    
    def define_matrix_detection(self,iline,monitor,nn):
        if iline[22]=='I' or iline[22]=='IVB' or iline[22]=='IIB' or iline[22]=='VI':
            MP,MC=self.set_matrix_detectiontypeI(iline,monitor,nn)
        else:
            MP,MC=None,None
        return MP,MC
            
    def analysis_from_assignednuclides(self,total_assigned_peaklist,monitor):
        self.quantification=[]
        nn=0
        for i in total_assigned_peaklist:
            if i==[] or i==None:
                self.quantification.append(None)
            else:
                spectrum_x=[]
                for th in i:
                    Mtrx=self.define_matrix(th,monitor,nn)
                    spectrum_x.append(Mtrx)
                self.quantification.append(spectrum_x)
            nn+=1
            
    def analysis_from_nuclidelist(self,monitor,nuclide_list,i,tolerance):
        nuclide_quantified_list=[]
        nuclide_detection_list=[]
        quantifiedMatrix_list=[]
        detectionMatrix_list=[]
        for ok in nuclide_list:
            nuclide_detection_list.append(ok)
            MD=self.define_matrix_detection(ok,monitor,i)
            detectionMatrix_list.append(MD)
        return nuclide_quantified_list,nuclide_detection_list,quantifiedMatrix_list,detectionMatrix_list

def openhyperlabfile(file):
    with open(file, newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
        S=[]
        for row in spamreader:
            S.append(row)
        S.pop(0)
        return S
    
def read_rptfile(file):
    def takeSecond(elem):
        return float(elem[4])
    statslim=40
    with open(file, "r") as f:
        data=f.readlines()
        for i in range(len(data)):
            data[i]=data[i].replace('\r\n','')
            data[i]=data[i].replace('\n','')
            if '*' in data[i]:
                data[i]=data[i].replace(' ','')
                if '*UNIDENTIFIEDPEAKSUMMARY*' in data[i]:
                    idx=i
                elif '*IDENTIFIEDPEAKSUMMARY*' in data[i]:
                    ids=i
            if '\x00\x00\x00\x00\x00' in data[i]:
                data[i]=''
            if '\x0c' in data[i]:
                data[i]=''
            if 'Microsoft' in data[i]:
                data[i]=''
            if 'Centroid' in data[i]:
                data[i]=''
            if 'Channel' in data[i]:
                data[i]=''
            if 'ORTEC' in data[i]:
                data[i]=''
            if 'Page' in data[i]:
                data[i]=''
            if 'Zero offset:' in data[i]:
                Z=str.split(data[i])[-2]
                try:
                    Z=float(Z)
                except:
                    Z=0
            if 'Gain:' in data[i]:
                G=str.split(data[i])[-2]
                try:
                    G=float(G)
                except:
                    G=1000000
            if 'Quadratic:' in data[i]:
                Q=str.split(data[i])[-2]
                try:
                    Q=float(Q)
                except:
                    Q=0
            if 'Spectrum' in data[i]:
                data[i]=''
    peaklist=[]
    while idx>-1:
        try:
            values=str.split(data[idx+4])
            if values!='' and values!=[]:
                float(values[0]),float(values[1]),float(values[2]),float(values[3]),float(values[4]),float(values[5]),float(values[6])
                if float(values[3])>0 and float(values[5])>0 and float(values[5])<statslim:
                    fpc=str(float(values[3])*float(values[5])/100)
                    FW=(float(values[6])-Z)/G
                    #print(values)
                    peaklist.append(['','','','',values[0],'',values[1],'',values[3],fpc,str(FW)[:4],'','','','','','','','','',''])
        except:
            break
        else:
            idx+=1
    while ids>-1:
        try:
            values=str.split(data[ids+4])
            if values!='' and values!=[]:
                float(values[1]),float(values[2]),float(values[3]),float(values[4]),float(values[5]),float(values[6]),float(values[7][:-1])
                if float(values[4])>0 and float(values[6])>0 and float(values[6])<statslim:
                    fpc=str(float(values[4])*float(values[6])/1000)
                    FW=(float(values[7][:-1])-Z)/G
                    peaklist.append(['','','','',values[1],'',values[2],'',values[4],fpc,str(FW)[:4],'','','','','','','','','',''])
        except:
            break
        else:
            ids+=1
    peaklist.sort(key=takeSecond)
    return peaklist

def read_rptfile2(file,statslim=40,set_forall=True):
    idx, ids = None, None
    def takeSecond(elem):
        return float(elem[4])
    def taketime(date,time):
        giorno, mese, anno = date.split('/')
        ore, minuti, secondi = time.split(':')
        dt = datetime.datetime(int(anno),int(mese),int(giorno),int(ore),int(minuti),int(secondi))
        return dt
    #statslim=40
    with open(file, "r") as f:
        data=f.readlines()
        for i in range(len(data)):
            data[i]=data[i].replace('\r\n','')
            data[i]=data[i].replace('\n','')
            if 'Start time:' in data[i]:
                strtime=data[i].split()[-2],data[i].split()[-1][:8]
                lvtime=int(data[i+1].split()[-1])
                rltime=int(data[i+2].split()[-1])
            if '*' in data[i]:
                data[i]=data[i].replace(' ','')
                if '*UNIDENTIFIEDPEAKSUMMARY*' in data[i]:
                    idx=i
                elif '*IDENTIFIEDPEAKSUMMARY*' in data[i]:
                    ids=i
            if '\x00\x00\x00\x00\x00' in data[i]:
                data[i]=''
            if '\x0c' in data[i]:
                data[i]=''
            if 'Microsoft' in data[i]:
                data[i]=''
            if 'Centroid' in data[i]:
                data[i]=''
            if 'Channel' in data[i]:
                data[i]=''
            if 'ORTEC' in data[i]:
                data[i]=''
            if 'Page' in data[i]:
                data[i]=''
            if 'Zero offset:' in data[i]:
                Z=str.split(data[i])[-2]
                try:
                    Z=float(Z.replace(',','.'))
                except:
                    Z=0
            if 'Gain:' in data[i]:
                G=str.split(data[i])[-2]
                try:
                    G=float(G.replace(',','.'))
                except:
                    G=1000000
            if 'Quadratic:' in data[i]:
                Q=str.split(data[i])[-2]
                try:
                    Q=float(Q)
                except:
                    Q=0
            if 'Spectrum' in data[i]:
                data[i]=''
    startcounting=taketime(*strtime)
    peaklist=[]
    if idx is not None:
        while idx>-1:
            try:
                values=str.split(data[idx+4])
                if values!='' and values!=[]:
                    float(values[0].replace(',','.')),float(values[1].replace(',','.')),float(values[2].replace(',','.')),float(values[3].replace(',','.')),float(values[4].replace(',','.')),float(values[5].replace(',','.')),float(values[6].replace(',','.'))
                    if float(values[3].replace(',','.'))>0 and float(values[5].replace(',','.'))>0 and float(values[5].replace(',','.'))<statslim:
                        fpc=str(float(values[3].replace(',','.'))*float(values[5].replace(',','.'))/100)
                        FW=(float(values[6].replace(',','.'))-Z)/G
                        peaklist.append(['','','','',values[0].replace(',','.'),'',values[1].replace(',','.'),'',values[3].replace(',','.'),fpc,format(FW,'.2f'),'','','','','','','','','',values[2].replace(',','.')])
            except:
                break
            else:
                idx+=1
    if ids is not None:
        while ids>-1:
            try:
                values=str.split(data[ids+4])
                if values!='' and values!=[]:
                    float(values[1].replace(',','.')),float(values[2].replace(',','.')),float(values[3].replace(',','.')),float(values[4].replace(',','.')),float(values[5].replace(',','.')),float(values[6].replace(',','.')),float(values[7][:-1].replace(',','.'))
                    if float(values[4].replace(',','.'))>0 and float(values[6].replace(',','.'))>0 and float(values[6].replace(',','.'))<statslim:
                        fpc=str(float(values[4].replace(',','.'))*float(values[6].replace(',','.'))/100)
                        FW=(float(values[7][:-1].replace(',','.'))-Z)/G
                        peaklist.append(['','','','',values[1].replace(',','.'),'',values[2].replace(',','.'),'',values[4].replace(',','.'),fpc,format(FW,'.2f'),'','','','','','','','','',values[3].replace(',','.')])
                    else:
                        if set_forall==False:
                            fpc=str(float(values[4].replace(',','.'))*float(values[6].replace(',','.'))/100)
                            FW=(float(values[7][:-1].replace(',','.'))-Z)/G
                            peaklist.append(['','','','',values[1].replace(',','.'),'',values[2].replace(',','.'),'',0.1,fpc,format(FW,'.2f'),'','','','','','','','','',values[3].replace(',','.')])
            except:
                break
            else:
                ids+=1
    peaklist.sort(key=takeSecond)
    return startcounting,rltime,lvtime,peaklist
    
def acquisiscispettroASC(n):
    f=open(n,'r')
    rl=f.readlines()
    f.close()
    r=[]
    for i in rl:
        r.append(i.replace('\n',''))
    a=0
    while a>-1:
        try:
            if '#LiveTime=' in r[a]:
                live=float(r[a].replace('#LiveTime=',''))
                afine=a-1
            if '#TrueTime=' in r[a]:
                real=float(r[a].replace('#TrueTime=',''))
            if '#AcqStart=' in r[a]:
                data=r[a].replace('#AcqStart=','')
                data=data.replace('T',' ')
                data,ora=str.split(data)
                anno,mese,giorno=str.split(data,'-')
                ora=ora.replace(':',' ')
                ore,minuti,secondi=str.split(ora)
            if '#LinEnergyCalParams=' in r[a]:
                linE=str.split(r[a].replace('#LinEnergyCalParams=',''))
                for l in range(len(linE)):
                    linE[l]=float(linE[l])
            if '#FwhmCalParams=' in r[a]:
                linW=str.split(r[a].replace('#FwhmCalParams=',''))
                for l in range(len(linW)):
                    linW[l]=float(linW[l])
        except IndexError:
            break
        else:
            a=a+1
    workinglist=r[:afine]
    spectrum_counts = [float(iks) for iks in workinglist]
    startcounting=datetime.datetime(int(anno),int(mese),int(giorno),int(ore),int(minuti),int(secondi))
    return startcounting,real,live,spectrum_counts,linE,linW

def read_chnfile2(n):
    import struct
    with open(n, "rb") as f:
        cs=4
        data = f.read(2)
        data = f.read(2)
        data = f.read(2)
        data = f.read(2)
        data = f.read(4)
        datev=struct.unpack('<I',data)
        real=int(datev[0])*20/1000
        data = f.read(4)
        datev=struct.unpack('<I',data)
        live=int(datev[0])*20/1000
        data = f.read(8)
        data = f.read(4)
        data = f.read(2)
        data = f.read(2)
        dtt=[]
        while data:
            data = f.read(cs)
            try:
                datev=struct.unpack('<I',data)
                datev=int(datev[0])
                dtt.append(datev)
            except:
                break
    lenit=65536
    if len(dtt)<65536:
        lenit=32768
    if len(dtt)<32768:
        lenit=16384
    if len(dtt)<16384:
        lenit=8192
    if len(dtt)<8192:
        lenit=4096
    dtt,rest=dtt[:lenit],dtt[lenit:]
    spectrum_counts = [int(iks) for iks in dtt]
    linE,linW=[0.0,0.0],[0.0,0.0]
    return real,live,spectrum_counts

def read_chnfile(n):
    import struct
    with open(n, "rb") as f:
        cs=4
        data = f.read(2)
        data = f.read(2)
        data = f.read(2)
        data = f.read(2)
        secs = str(data,'utf-8')
        data = f.read(4)
        datev=struct.unpack('<I',data)
        real=int(datev[0])*20/1000
        data = f.read(4)
        datev=struct.unpack('<I',data)
        live=int(datev[0])*20/1000
        data = f.read(8)
        ddmmmyy = str(data,'utf-8')
        ddmmmyy,y1 = ddmmmyy[:-1],ddmmmyy[-1:]
        if y1=='1':
            y2000=2000
        else:
            y2000=0
        day,month,year=ddmmmyy[:2],ddmmmyy[2:5],int(ddmmmyy[5:])+y2000
        monthasint={'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06','Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}
        month=monthasint.get(month)
        data = f.read(4)
        hhmm = str(data,'utf-8')
        data = f.read(2)
        data = f.read(2)
        startcounting=datetime.datetime(int(year),int(month),int(day),int(hhmm[:2]),int(hhmm[2:]),int(secs))
        dtt=[]
        while data:
            data = f.read(cs)
            try:
                datev=struct.unpack('<I',data)
                datev=int(datev[0])
                dtt.append(datev)
            except:
                break
    lenit=65536
    if len(dtt)<65536:
        lenit=32768
    if len(dtt)<32768:
        lenit=16384
    if len(dtt)<16384:
        lenit=8192
    if len(dtt)<8192:
        lenit=4096
    dtt,rest=dtt[:lenit],dtt[lenit:]
    spectrum_counts = [float(iks) for iks in dtt]
    linE,linW=[0.0,0.0],[0.0,0.0]
    return startcounting,real,live,spectrum_counts,linE,linW

def searchforalternateopenfile(unclimit=40,set_forall=True,title=None):
    """return nomeHyperLabfile,startcounting,realT,liveT,peak_list,spectrum_counts,linE,linW"""
    types=[('rpt file','.rpt')]
    nomeHyperLabfile=askopenfilename(title=title, filetypes=types)
    try:
        startcounting,realT,liveT,peak_list=read_rptfile2(nomeHyperLabfile,unclimit,set_forall)
    except:
        print(f'failed to import rpt: {nomeHyperLabfile}')
        return None,None,None,None,None,None,None,None
    else:
        nomeHyperLabfile=nomeHyperLabfile.replace(nomeHyperLabfile[-4:],'.chn')
        try:
            realT,liveT,spectrum_counts=read_chnfile2(nomeHyperLabfile)
        except:
            spectrum_counts = None#np.array([0]*8192)
            print(f'failed to import chn: {nomeHyperLabfile} - spectrum profile not available')
        return nomeHyperLabfile,startcounting,realT,liveT,peak_list,spectrum_counts,None,None

def searchforhypelabfile():
    types=[('csv file','.csv'),('rpt file','.rpt')]
    nomeHyperLabfile=askopenfilename(filetypes=types)
    if nomeHyperLabfile!=None and nomeHyperLabfile!='':
        if nomeHyperLabfile[-4:].lower() == '.csv':
            try:
                peak_list=openhyperlabfile(nomeHyperLabfile)
            except:
                peak_list=[]
        else:
            try:
                peak_list=read_rptfile(nomeHyperLabfile)
            except:
                peak_list=[]
        try:
            peak_list[0][20]
        except IndexError:
            startcounting,realT,liveT,spectrum_counts,linE,linW,peak_list=None,None,None,None,None,None,None
            print(f'failed to import peak list: {nomeHyperLabfile}')
        else:
            nomeHyperLabfile=nomeHyperLabfile.replace(nomeHyperLabfile[-4:],'.ASC')
            if os.path.isfile(nomeHyperLabfile):
                try:
                    startcounting,realT,liveT,spectrum_counts,linE,linW=acquisiscispettroASC(nomeHyperLabfile)
                except:
                    startcounting,realT,liveT,spectrum_counts,linE,linW=None,None,None,None,None,None
                    print(f'failed to import spectrum: {nomeHyperLabfile}')
            else:
                nomeHyperLabfile=nomeHyperLabfile.replace(nomeHyperLabfile[-4:],'.chn')
                try:
                    startcounting,realT,liveT,spectrum_counts,linE,linW=read_chnfile(nomeHyperLabfile)
                except:
                    startcounting,realT,liveT,spectrum_counts,linE,linW=None,None,None,None,None,None
                    print(f'failed to import spectrum: {nomeHyperLabfile}')
    else:
        startcounting,realT,liveT,spectrum_counts,linE,linW,peak_list=None,None,None,None,None,None,None
    return nomeHyperLabfile,startcounting,realT,liveT,peak_list,spectrum_counts,linE,linW

def searchrptfilesforcalibration(unclimit=40,set_forall=True):
    types=[('rpt file','.rpt')]
    nomeHyperLabfiles=askopenfilenames(filetypes=types)
    AAA=[]
    if nomeHyperLabfiles!=None and nomeHyperLabfiles!='':
        for nomeHyperLabfile in nomeHyperLabfiles:
            try:
                startcounting,realT,liveT,peak_list=read_rptfile2(nomeHyperLabfile,unclimit,set_forall)
            except:
                peak_list=[]
            try:
                peak_list[0][20]
            except IndexError:
                print(f'failed to import peak list: {nomeHyperLabfile}')
            else:
                AAA.append([nomeHyperLabfile,startcounting,realT,liveT,peak_list])
    return AAA

def searchforalternateopenmultiplefiles(unclimit=40,set_forall=True, title=None):
    types=[('rpt file','.rpt')]
    nomeHyperLabfiles=askopenfilenames(title=title, filetypes=types)
    AAA=[]
    if nomeHyperLabfiles!=None and nomeHyperLabfiles!='':
        for nomeHyperLabfile in nomeHyperLabfiles:
            try:
                startcounting,realT,liveT,peak_list=read_rptfile2(nomeHyperLabfile,unclimit,set_forall)
            except:
                print(f'failed to import rpt: {nomeHyperLabfile}')
            else:
                nomeHyperLabfile=nomeHyperLabfile.replace(nomeHyperLabfile[-4:],'.chn')
                try:
                    realT,liveT,spectrum_counts=read_chnfile2(nomeHyperLabfile)
                except:
                    spectrum_counts = None#np.array([0]*8192)
                    print(f'failed to import chn: {nomeHyperLabfile} - spectrum profile not available')
                AAA.append([nomeHyperLabfile,startcounting,realT,liveT,peak_list,spectrum_counts,None,None])
    return AAA

def searchforhypelabmultiplefiles():
    types=[('csv file','.csv'),('rpt file','.rpt')]
    nomeHyperLabfiles=askopenfilenames(filetypes=types)
    AAA=[]
    if nomeHyperLabfiles!=None and nomeHyperLabfiles!='':
        for nomeHyperLabfile in nomeHyperLabfiles:
            if nomeHyperLabfile[-4:].lower() == '.csv':
                try:
                    peak_list=openhyperlabfile(nomeHyperLabfile)
                except:
                    peak_list=[]
            else:
                try:
                    peak_list=read_rptfile(nomeHyperLabfile)
                except:
                    peak_list=[]
            try:
                peak_list[0][20]
            except IndexError:
                startcounting,realT,liveT,spectrum_counts,linE,linW,peak_list=None,None,None,None,None,None,None
                print(f'failed to import peak list: {nomeHyperLabfile}')
            else:
                nomeHyperLabfile=nomeHyperLabfile.replace(nomeHyperLabfile[-4:],'.ASC')
                if os.path.isfile(nomeHyperLabfile):
                    try:
                        startcounting,realT,liveT,spectrum_counts,linE,linW=acquisiscispettroASC(nomeHyperLabfile)
                    except:
                        startcounting,realT,liveT,spectrum_counts,linE,linW=None,None,None,None,None,None
                        print(f'failed to import spectrum: {nomeHyperLabfile}')
                else:
                    nomeHyperLabfile=nomeHyperLabfile.replace(nomeHyperLabfile[-4:],'.chn')
                    try:
                        startcounting,realT,liveT,spectrum_counts,linE,linW=read_chnfile(nomeHyperLabfile)
                    except:
                        startcounting,realT,liveT,spectrum_counts,linE,linW=None,None,None,None,None,None
                        print(f'failed to import spectrum: {nomeHyperLabfile}')
                AAA.append([nomeHyperLabfile,startcounting,realT,liveT,peak_list,spectrum_counts,linE,linW])
    return AAA
        
if __name__=='__main__':
    searchforhypelabfile()
