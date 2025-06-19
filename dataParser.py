import pandas as pd
import warnings
import matplotlib.pyplot as plt
from docx.shared import Inches
import io
from report import *
import numpy as np

class Data():
    def download_from_sharepoint(self,site_url,file_url,new_filename,username,password):
        from office365.runtime.auth.authentication_context import AuthenticationContext
        from office365.sharepoint.client_context import ClientContext
        from office365.sharepoint.files.file import File    
        import os
        import tempfile
        ctx_auth = AuthenticationContext(site_url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url, ctx_auth)
            with open(new_filename, 'wb') as output_file:
                file = (
                        ctx.web.get_file_by_server_relative_url(file_url).download(output_file).execute_query()
                )
            print("[Ok] file has been downloaded into: {0}".format(new_filename))    
    
    def load_file(self, path: str, sheet: str):
        """Load data from an Excel file."""
        try:
            warnings.simplefilter(action='ignore', category=UserWarning)
            return pd.read_excel(path, sheet_name=sheet)
        except Exception as loadErr:
            print(loadErr)
            return pd.array()
        
class ParsingData():
    def __init__(self,week,lgbk_sh,discrepancy,preventive_maintenance,rtms,year):
        self.week=week
        self.year=year
        self.lgbksh=lgbk_sh
        self.discrepancy=discrepancy
        self.preventive_maintenance=preventive_maintenance
        self.rtms=rtms
        

    def parse_logbook_sh(self):
        output={}
        deviation_details=[]
        outcome_out_range=["DCO","SDC","ERR"]
        for i,row in self.lgbksh[10:].iterrows():
            if str(row.iloc[2])==str(self.week) and str(row.iloc[1]).split('-')[0] == str(self.year):
                if str(row.iloc[22]) not in outcome_out_range:
                    notes=""
                    if str(row.iloc[24])=="nan": notes=str(row.iloc[63])
                    elif str(row.iloc[63])=="nan": notes=str(row.iloc[24])
                    else: notes=str(row.iloc[24])+"\n"+str(row.iloc[63])
                    deviation_details.append(
                        (str(row.iloc[69]),  #AJT ID
                         str(row.iloc[1].date()),#.split(" ")[0], #Date
                         str(row.iloc[10]),#Device 
                         #str(row.iloc[22]),#Outcome,
                         str(row.iloc[71]),#Original Outcome,
                         str(row.iloc[23]),#Deviation  
                         notes,#Notes 
                         
                        )  
                    )
                if str(row.iloc[1]).split(" ")[0] in output:
                    output[str(row.iloc[1]).split(" ")[0]].append(
                                                    (str(row.iloc[10]),
                                                    str(row.iloc[71]),
                                                    str(row.iloc[23])
                                                    ))
                else: output[str(row.iloc[1]).split(" ")[0]]=[(str(row.iloc[10]),
                                                    str(row.iloc[71]),
                                                    str(row.iloc[23])
                                                    )]
        return (output,deviation_details)
    
    
    
    def simulator_utilization_data(self,output:dict,dates:list):
        sim_util_table={}
        devices=['FMS1','FMS2','PTT1','PTT2','PTT3','ULTD1','ULTD2','LVC']
        for date in dates:
            tmp=[[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0]] #FMS1:(planned,completed) FMS2:(planned,completed) PTT1:(planned,completed) PTT2:(planned,completed) PTT3:(planned,completed) 
                                                                      #ULTD1:(planned,completed) ULTD2:(planned,completed) LVC:(planned,completed)
            if str(date) in output:
                training=output[str(date)]
                for x in training:
                    pos=devices.index(x[0])
                    if x[1] !='RSLD' and x[1] !='ERR':
                        tmp[pos][0]=tmp[pos][0]+1
                    if x[1] =='DCO' or x[1] =='SDC':
                        tmp[pos][1]=tmp[pos][1]+1
                sim_util_table[date]=tmp
            else: sim_util_table[date]=[[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0]]
        return sim_util_table

    def issues_a_b_c_d_na(self):
        output={'FMS':[],'PTT':[],'ULTD':[],'SBT':[],'CBT':[],'MPDS':[]}
        
        for i,row in self.discrepancy[2:].iterrows():
                #if str(row['Week']) == str(self.week) and str(row['Year'])== str(self.year):
            if row['Date'].date() in get_dates_in_week(year=self.year,week=self.week):
                family_sim=''.join([c for c in str(row['Device']) if c.isalpha()])
                sbt_family=('AllSBTs','SBT Room1','SBT Room2','SBTRoom')
                cbt_family=('CBT Room1','CBT Room2')
                if family_sim in sbt_family: family_sim='SBT'
                elif family_sim in cbt_family: family_sim='CBT'
                output[family_sim].append((
                    str(row['ID']), 
                    str(row['Device']),
                    str(row['Date'].date()),#.split(" ")[0],
                    str(row['Severity']),
                    str(row['Observation']),
                    str(row['Action/Comment/Workaround']),
                    str(row['Discrepancy Status'])
                    ))
                
        #print(output)
        return output
    
    def preventiveM(self,days=[]):
        output=[]
        for i,row in self.preventive_maintenance[2:].iterrows():
            if row['Date'].date() in days:
                output.append((
                    str(row['Task id']),
                    str(row['Date'].date()),
                    str(row['Device']),
                    str(row['Task Description']),
                    str(row['Period'])
                ))
        return output
    
    def rtms_log(self,days=[]):
        output=[]
        for i,row in self.rtms[6:].iterrows():
            try:
                if row.iloc[1].date() in days:
                    output.append((
                        
                        str(row.iloc[0]), #ID
                        str(row.iloc[1].date()), #Date
                        str(row.iloc[5]).replace('nan','N/A'), #mission
                        str(row.iloc[6]).replace('nan','N/A'),#area
                        str(row.iloc[7]).replace('nan','N/A'),#A/C
                        str(row.iloc[8]).replace('nan','N/A'),#POD ID
                        str(row.iloc[9]).replace('nan','N/A'),#DL freq
                        str(row.iloc[10]).replace('nan','N/A'),#radio ch
                        f"COMMENTS RTMP IP: {str(row.iloc[12]).replace('nan','N/A')}\nCOMMENTS AC IP: {str(row.iloc[14]).replace('nan','N/A')}\nMPDS OPERATOR  COMMENTS: {str(row.iloc[15]).replace('nan','N/A')}"
                    ))
            except: pass
        return output
    
    def chart_daily_discrepancies_data(self,days=[]):
        lgbk_sh=self.parse_logbook_sh()[0] #dictionary
        tmp_discrepancy=self.issues_a_b_c_d_na()#dictionary
        output={}
        for day in days:
            output[str(day)]={'DNCO':0,'A':0,'B':0,'C':0,'D':0,'nan':0}
            try:
                for session_day in lgbk_sh[str(day)]:
                    if session_day[1]=='DNCO': #or session_day[1]=='SDNC':
                        output[str(day)]['DNCO']+=1
            except: pass
            
        for fam in ['FMS','PTT','ULTD','SBT','CBT']:
            tmp=tmp_discrepancy[fam]
            for discrepancy in tmp:
                severity=discrepancy[3]
                output[discrepancy[2]][severity]+=1
        for keys in output.keys():
            output[keys]['N/A']=output[keys]['nan']
            del output[keys]['nan']
  
        #print(output)
        return output
    
    def   chart_weekly_discrepancies_data(self,sim_util:dict):
        discrepancies=self.issues_a_b_c_d_na()
        sim_fam={'FMS':0,'PTT':0,'ULTD':0}#,'SBT':0,'CBT':0} #SBT and CBT excluded just to compare the number of discrepancies with the number of scheduled sessions
        sim_util_fam={'FMS':0,
                      'PTT':0,
                      'ULTD':0
                      }
        for days in list(sim_util.keys()):
            sim_util_fam['FMS']+=sim_util[days][0][0]+sim_util[days][1][0]
            sim_util_fam['PTT']+=sim_util[days][2][0]+sim_util[days][3][0]+sim_util[days][4][0]
            sim_util_fam['ULTD']+=sim_util[days][5][0]+sim_util[days][6][0]
        
        #The formula to calculate the normalized discrepancies for every family sim is:
        # [Total_Discr/Total_Sess]/[Family_Sim_Discr/Family_Sim_Sess]
        
        total_discr=len(discrepancies['FMS']) + len(discrepancies['PTT']) + len(discrepancies['ULTD']) 
        total_sess=sim_util_fam['FMS']+sim_util_fam['PTT']+sim_util_fam['ULTD']
        
        for fam in sim_fam.keys():
            if sim_util_fam[fam]==0:
                sim_fam[fam]=1
            else: sim_fam[fam]=(total_discr/total_sess)*(len(discrepancies[fam])/sim_util_fam[fam])
        
        return sim_fam,discrepancies

    

    def generate_weekly_data_integer_values(self,title:str,sim_util:dict):
        data=self.chart_weekly_discrepancies_data(sim_util)
        labels=list(data.keys())
        values=list(data.values())
        max_y=int(max(values) / 2) *2 +2
        plt.ylim(0,max_y+max_y*0.1)
        plt.figure(figsize=(8,6))
        bars=plt.bar(labels,values,color='steelblue',width=0.5)
        plt.title(title)
        # Aggiungi etichette e legenda
        plt.yticks(np.arange(0, max_y+1, 2))
        plt.grid(axis='y', linestyle='--', alpha=0.7,linewidth=0.2)
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, yval + 0.2, int(yval), ha='center', va='bottom')
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', bbox_inches='tight')
        plt.close()
        img_stream.seek(0)
        return img_stream
    
    def generate_weekly_data(self, title: str, sim_util: dict):
        import math
        import numpy as np
        import matplotlib.pyplot as plt
        import matplotlib.ticker as mticker
        import io

        data, discrepancies = self.chart_weekly_discrepancies_data(sim_util)
        #data to plot
        labels = list(data.keys())
        labels.append('SBT')
        labels.append('CBT')
        values_discr_perc = list(data.values())
        values_discr_perc.append(0)
        values_discr_perc.append(0)
        total=sum(values_discr_perc)
        values_discr_perc = [(x/total)*100 for x in values_discr_perc]
        
        values_discr_abs=[len(discrepancies[fam]) for fam in labels]

        x = np.arange(len(labels))
        bar_width = 0.9
        fig, ax1 = plt.subplots(figsize=(18, 15))

        # Primo asse (sinistra, percentuali)
        bars1=ax1.bar(x, values_discr_perc, width=bar_width, color="#1676D1", label="Normalized Issue Rate %", alpha=0.2, zorder=0)
        ax1.set_ylabel("Normalized Issue Rate %", fontsize=16)
        #ax1.set_ylim(0, max(values_discr_perc)*1.2)
        #ax1.set_yticks(np.linspace(0, max(values_discr_perc)*1.2 + 1, 10))
        ax1.set_ylim(0, 100)
        #ax1.set_yticks(np.linspace(0, 100, 10))
        ax1.set_yticks(range(0, 101, 10))
        ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{int(x)}%" ))
        ax1.tick_params(axis='y', labelsize=16)
        ax1.grid(True, axis='y', linestyle='--', alpha=0.5, zorder=0)

        # Secondo asse (destra, valori assoluti)
        ax2 = ax1.twinx()
        bars2=ax2.bar(x, values_discr_abs, width=bar_width/2, color="#2E3E6B96", label="Issues", alpha=1, zorder=3)
        ax2.set_ylabel("Issues", fontsize=16)
        ax2.set_ylim(0, max(values_discr_abs)*1.8)
        ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{int(x)}"))
        ax2.tick_params(axis='y', labelsize=16)

        # Valori sopra le barre percentuali
        for bar in bars1:
            height = bar.get_height()
            if height > 0:
                ax1.text(bar.get_x() + bar.get_width()*0.1, height + 1,
                        f"{height:.1f}%", ha='center', va='bottom', fontsize=16)
        # Valori sopra le barre assolute
        for bar in bars2:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()*0.5, height,
                    f"{int(height)}", ha='center', va='bottom', fontsize=16)

        # Etichette e legende
        ax1.set_xticks(x)
        ax1.set_xticklabels(labels, fontsize=16)

        # Legenda combinata
        handles1, labels1 = ax1.get_legend_handles_labels()
        handles2, labels2 = ax2.get_legend_handles_labels()

        # Combina
        all_handles = handles1 + handles2
        all_labels = labels1 + labels2

        # Legenda combinata in basso centrata
        plt.legend(all_handles, all_labels,
                loc='lower center',
                bbox_to_anchor=(0.5, -0.12),
                ncol=3,
                fontsize=18)

        img_stream = io.BytesIO()
        plt.tight_layout()
        plt.savefig(img_stream, format='png', bbox_inches='tight')
        plt.close()
        img_stream.seek(0)
        return img_stream
    
    def generate_weekly_data_old(self, title: str, sim_util: dict):
        import math
        import numpy as np
        import matplotlib.pyplot as plt
        import io

        data, discrepancies = self.chart_weekly_discrepancies_data(sim_util)
        print(data)
        labels = list(data.keys())[:3]
        values_discr_perc = list(data.values())[:3]
        values_discr_abs = [math.log10(len(discrepancies[fam])) if len(discrepancies[fam]) > 0 else 0 for fam in labels]
        tmp = sum(len(discrepancies[fam]) for fam in labels)
        total_discr_abs = np.full(len(labels), math.log10(tmp))

        total = sum(values_discr_perc)
        values_pct = [(v / total) * 100 if total > 0 else 0 for v in values_discr_perc]

        x = np.arange(len(labels))
        bar_width = 0.35

        fig, ax1 = plt.subplots(figsize=(8, 6))
        ax1.set_ylabel("Discrepancies %", color='steelblue')
        ax1.set_ylim(0, 105)
        ax1.set_xticks(x)
        ax1.set_xticklabels(labels)

        # === Barre percentuali (asse sinistro) ===
        bars1 = ax1.bar(x - bar_width/2, values_pct, width=bar_width, color='steelblue', label="Discrepancies %")
        for bar in bars1:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2, height + 1, f"{height:.1f}%", ha='center', va='bottom', fontsize=8)

        ax1.tick_params(axis='y', labelcolor='steelblue')
        ax1.grid(axis='y', linestyle='--', alpha=0.7, linewidth=0.3)

        # === Asse destro ===
        ax2 = ax1.twinx()
        ax2.set_ylabel("Discrepancies Abs", color='orange')
        ax2.set_ylim(0, max(total_discr_abs) * 1.2)

       

        # Barre arancioni (log absolute discrepancies)
        bars2 = ax2.bar(x + bar_width/2, values_discr_abs, width=bar_width, color='orange', label="Discrepancies Abs")
        """for bar in bars2:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2, height + 0.1, f"{height:.1f}", ha='center', va='bottom', fontsize=8)"""

        ax2.tick_params(axis='y', labelcolor='orange')

        # Titolo e legende
        plt.title(title)
        handles1, labels1 = ax1.get_legend_handles_labels()
        handles2, labels2 = ax2.get_legend_handles_labels()

        # Combina
        all_handles = handles1 + handles2
        all_labels = labels1 + labels2

        # Legenda combinata in basso centrata
        plt.legend(all_handles, all_labels,
                loc='lower center',
                bbox_to_anchor=(0.5, -0.12),
                ncol=3,
                fontsize=10)

        # Salva immagine
        img_stream = io.BytesIO()
        plt.tight_layout()
        plt.savefig(img_stream, format='png', bbox_inches='tight')
        plt.close()
        img_stream.seek(0)
        return img_stream




    
    def generate_daily_data(self,title:str,days:list):
        data=self.chart_daily_discrepancies_data(days)
        A=[]
        B=[]
        C=[]
        D=[]
        NA=[]
        DNCO=[]
        total=[]
        for k in data.keys():
            A.append(data[k]['A'])
            B.append(data[k]['B'])
            C.append(data[k]['C'])
            D.append(data[k]['D'])
            NA.append(data[k]['N/A'])
            DNCO.append(data[k]['DNCO'])
            total.append(data[k]['A']+data[k]['B']+data[k]['C']+data[k]['D']+data[k]['N/A'])
        x = np.arange(len(days))
        bar_width = 0.9
        group_bar_width = bar_width  / 5
        offsets = np.linspace(-group_bar_width*2, group_bar_width*2, 5)
        plt.figure(figsize=(18, 15))
        plt.bar(x, total, color="#9DC3E6", label="Total", alpha=0.4,zorder=0, width=bar_width)
        plt.bar(x, DNCO, bottom=total, color="#FF0000", label="DNCO",alpha=0.6, zorder=0, width=bar_width)
        max_y=int(max(total[i]+DNCO[i] for i in range(5))/2)*2 +2
        plt.ylim(0,max_y)
        for i in range(len(x)):   
            #if A[i] >0:
                #plt.text(x[i] + offsets[0], 0.05, str(A[i]), ha='center', va='bottom', fontsize=16, color='gray')
            #if B[i] >0:
                #plt.text(x[i] + offsets[1], 0.05, str(B[i]), ha='center', va='bottom', fontsize=16, color='gray')
            #if C[i] >0:
                #plt.text(x[i] + offsets[2], 0.05, str(C[i]), ha='center', va='bottom', fontsize=16, color='gray')
            #if D[i] >0:
                #plt.text(x[i] + offsets[3], 0.05, str(D[i]), ha='center', va='bottom', fontsize=16, color='gray')
            #if NA[i] >0:
                #plt.text(x[i] + offsets[4], 0.05, str(NA[i]), ha='center', va='bottom', fontsize=16, color='gray')
        
            if total[i] >0:
                plt.text(x[i], total[i]/2, str(total[i]), ha='center', va='center', fontsize=16, color='black')# label inside Total    
            if DNCO[i] >0:
                plt.text(x[i], total[i] + DNCO[i]/2, str(DNCO[i]), ha='center', va='center', fontsize=16, color='black')# label inside DNCO 

        # Disegna le barre A–N/A in primo piano (sopra le stacked)
        plt.bar(x + offsets[0], A, width=group_bar_width, color="#A95720", label='A', zorder=2)
        plt.bar(x + offsets[1], B, width=group_bar_width, color="#C46627", label='B', zorder=2)
        plt.bar(x + offsets[2], C, width=group_bar_width, color="#ED7D31", label='C', zorder=2)
        plt.bar(x + offsets[3], D, width=group_bar_width, color="#F09E7A", label='D', zorder=2)
        plt.bar(x + offsets[4], NA, width=group_bar_width, color="#D6D6D6", label='N/A', zorder=2) #gray

        # Aggiungi etichette e legenda
        
        plt.yticks(np.arange(0, max_y+1, 2),fontsize=16)
        plt.xticks(x, days,fontsize=16)
        plt.title(title,fontsize=20)
        plt.legend(loc='lower center', bbox_to_anchor=(0.5, -0.12),ncol=7,fontsize=16)
        plt.grid(True, axis='y', linestyle='--', alpha=0.5)

        # Salva in buffer per Word
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', bbox_inches='tight')
        plt.close()
        img_stream.seek(0)
        return img_stream