from dataParser import *
from report import *
import os
from datetime import date
try:
    from dotenv import dotenv_values, load_dotenv
except:
    from dotenv import main 
import sys


local_path=f"{str(os.getcwd())}\\"
try:
    env=dotenv_values(local_path+'env.env')
    logbooksh_url_path=env.get('logbooksh_url')
    limitation_url_path=env.get('limitation_url')
    rtmslog_url_path=env.get('rtmslog_url')
    site_url_path=env.get('site_url')
    dailylog_url_path=env.get('dailylog_url')
    username=env.get('username')
    password=env.get('password')
    
except: 
    env=main.load_dotenv(local_path+'env.env')
    logbooksh_url_path=os.getenv('logbooksh_url')
    limitation_url_path=os.getenv('limitation_url')
    rtmslog_url_path=os.getenv('rtmslog_url')
    site_url_path=os.getenv('site_url')
    dailylog_url_path=os.getenv('dailylog_url')
    username=os.getenv('username')
    password=os.getenv('password')
    

#START DOWNLOADING DOCUMENTS FROM SHAREPOINT
data_container=Data()
data_container.download_from_sharepoint(site_url_path,logbooksh_url_path,local_path+"Record of SH Duty Exercise and Times Log_Total.xlsm",username,password)
#data_container.download_from_sharepoint(site_url_path,limitation_url_path,local_path+"Limitation Logbook.xlsm",username,password)
data_container.download_from_sharepoint(site_url_path,rtmslog_url_path,local_path+"RTMS_LOGBOOK_V1.xlsx",username,password)
data_container.download_from_sharepoint(site_url_path,dailylog_url_path,local_path+"LogBookEventIssue.xlsm",username,password)
#STOP DOWNLOADING DOCUMENTS FROM SHAREPOINT""

#START TO READ THE DOCUMENTS
logbook_sh=data_container.load_file(local_path+"Record of SH Duty Exercise and Times Log_Total.xlsm","Log Book")
#limitation=data_container.load_file(local_path+"Limitation Logbook.xlsm","Limitation Log")
discrepancy=data_container.load_file(local_path+"LogBookEventIssue.xlsm","Discrepancy")
preventive_maintenance=data_container.load_file(local_path+"LogBookEventIssue.xlsm","PM")
rtms=data_container.load_file(local_path+"RTMS_LOGBOOK_V1.xlsx","RTMS LOGBOOK")
#STOP TO READ THE DOCUMENTS


#START TO CANCEL THE DOCUMENTS FROM THE LOCAL PATH
"""os.remove(local_path+"Record of SH Duty Exercise and Times Log_Total.xlsm")
os.remove(local_path+"RTMS_LOGBOOK_V1.xlsx")
os.remove(local_path+"LogBookEventIssue.xlsm")"""
#os.remove(local_path+"Limitation Logbook.xlsm")
#STOP TO CANCEL THE DOCUMENTS FROM THE LOCAL PATH


def create(logbook_sh,discrepancy,preventive_maintenance,rtms,week,year=2025):

    #START READING DATA
    pdata=ParsingData(week,logbook_sh,discrepancy,preventive_maintenance,rtms,year=year)
    output,deviation_details=pdata.parse_logbook_sh()
    sim_util=pdata.simulator_utilization_data(output,get_dates_in_week(year,week))
    preventive=pdata.preventiveM(get_dates_in_week(year,week))
    rtmsLog=pdata.rtms_log(get_dates_in_week(year,week))
    if not output and len(deviation_details)==0 and not sim_util:
        print(f'No data available for week {week} of year {year}!')
        return False
    #STOP READING DATA

    #START REPORT GENERATION
    r=Report(local_path,year,week)
    r.doc.add_page_break()
    #TABLE 1
    r.new_paragraph("Simulator Utilization")
    r.generate_text('Below table provide an overview of the planned vs performed sessions per device during the reported week.')
    r.decorate_table(r.generate_sim_util_table(sim_util),   header_rows=2,
                                                            table_alignment_center=False,
                                                            header_bg='BDD7EE',   # light blue
                                                            font_name='Calibri',
                                                            font_size=8,
                                                            bold_headers=True,
                                                            total_width_cm=20.0,
                                                            column_widths_cm=[1.4, 1.86, 1.86, 1.86, 1.86, 1.86, 1.86, 1.86, 1.86, 1.86, 1.86])
    #TABLE 2
    r.new_paragraph("Deviation Details")
    r.generate_text('Below table provide details of deviations that occurred during training sessions with an abnormal outcome during the reported week.')
    if len(deviation_details)==0:
        r.generate_text("All sessions have been completed successfully.")
    else:
        r.decorate_table( r.generate_generic_table(6,deviation_details,['S/N','Date','Device','Outcome','Deviation','Notes']),
                                                                header_rows=1,
                                                                table_alignment_center=False,
                                                                header_bg='BDD7EE',   # light blue
                                                                font_name='Calibri',
                                                                font_size=9,
                                                                bold_headers=True,
                                                                total_width_cm=21.2,
                                                                column_widths_cm=[2.1, 1.9, 1.7, 1.7, 1.7, 11.6],
                                                                columns_left_alignment=[5])
                                                            
   

    r.decorate_table(                                       r.legend('(REASON FOR) CANCELATION DEVIATION',rows=6,cols=4,cell_content=[('CD1','No Show user','CD6','Device under service'),
                                                                                                                                      ('CD2','No show Instructor','CD7','Facility not ready'),
                                                                                                                                      ('CD3','Device failed during session','CD8','Software Development'),
                                                                                                                                      ('CD4','Device not ready for training','CD9','Hardware Development'),
                                                                                                                                      ('CD5','2nd SIM not available for NETMODE','CD10','External Factors')]),
                                                            table_alignment_center=False,
                                                            header_rows=1,
                                                            header_bg='BDD7EE',   # light blue
                                                            font_name='Calibri',
                                                            font_size=9,
                                                            bold_headers=True,
                                                            align_center=False,
                                                            total_width_cm=14.6,
                                                            column_widths_cm=[1.3,6,1.3,6],
                                                            alternate_row_color='FFFFFF')
    #CHART 1
    r.doc.add_page_break()
    r.new_paragraph("Weekly Discrepancies per Simulator")
    r.generate_text('Below graph provides an overview of all discrepancies per device during the reported week. Please refer to ANNEX A for the discrepancy Reports.')
    r.add_chart(pdata.generate_weekly_data("Discrepancies A - B - C- D - N/A",sim_util))
    #CHART 2
    r.doc.add_page_break()
    r.new_paragraph("Daily Discrepancies Severities")
    r.generate_text('Below graph provides a per day overview of the severity levels for the discrepancies and how many DNCO occurred during the reported week. Please refer to ANNEX A for the discrepancy Reports.')
    r.add_chart(pdata.generate_daily_data(title="Severity A - B - C - D - N/A VS DNCO",days=get_dates_in_week(year,week)))
    r.decorate_table(                                       r.legend('Legend',rows=7,cols=2,cell_content=[('A','Discrepancy with major impact on training'),
                                                                                                          ('B','Discrepancy with moderate impact on training'),
                                                                                                          ('C','Discrepancy with minor impact on training'),
                                                                                                          ('D','Discrepancy with no impact or not-technical'),
                                                                                                          ('N/A','Discrepancy that occurred outside a training session'),
                                                                                                          ('DNCO','Duty Not Carried Out')]),
                                                            table_alignment_center=False,
                                                            header_rows=1,
                                                            header_bg='BDD7EE',   # light blue
                                                            font_name='Calibri',
                                                            font_size=9,
                                                            bold_headers=True,
                                                            align_center=False,
                                                            total_width_cm=9.3,
                                                            column_widths_cm=[1.3,8],
                                                            alternate_row_color='FFFFFF')

    #TABLE 3
    r.doc.add_page_break()
    r.new_paragraph("Executed Preventive Maintenance")
    r.generate_text('Below list provides an overview of all preventive maintenance tasks that were execute during the reported week.')
    r.decorate_table( r.generate_generic_table(5,preventive,['Task ID','Date','Device','Task Description','Period']),
                                                            header_rows=1,
                                                            table_alignment_center=False,
                                                            header_bg='BDD7EE',   # light blue
                                                            font_name='Calibri',
                                                            font_size=9,
                                                            bold_headers=True,
                                                            total_width_cm=20.0,
                                                            column_widths_cm=[3, 3, 3, 8, 3],
                                                            columns_left_alignment=[3])
    
    r.doc.add_page_break()
    #TABLE 4
    r.convert_in_landscape()
    r.new_paragraph("RTMS")
    r.generate_text('Below table provide an overview of the performed sessions on the RTMS system during the reported week.')
    if len(rtmsLog)==0:
        r.generate_text("No RTMS sessions found for this week.")
    else:
        r.decorate_table( r.generate_generic_table(9,rtmsLog,['ID','Date','Mission','Area','Aircraft','POD ID','DL Freq','Radio CH','Notes']),
                                                                table_alignment_center=False,
                                                                header_rows=1,
                                                                header_bg='BDD7EE',   # light blue
                                                                font_name='Calibri',
                                                                font_size=8,
                                                                bold_headers=True,
                                                                total_width_cm=25.0,
                                                                column_widths_cm=[2, 2, 2, 2, 2, 2, 2, 2, 9],
                                                                columns_left_alignment=[8]
                                                                )
    #ANNEX
    r.convert_in_portrait()
    r.new_paragraph('ANNEX A')
    r.generate_text('Discrepancy Reports created during the reported week.')
    r.save_documents()
    #STOP REPORT GENERATION

#create(logbook_sh,discrepancy,preventive_maintenance,rtms,week=21,year=2025)

def main__():
    try:
        if len(sys.argv) > 2:
            if sys.argv[1] == '--year':
                year = int(sys.argv[2])
            if sys.argv[3] == '--week':
                week = int(sys.argv[4]) 
        else:
            previous_week_date=date.today()-timedelta(weeks=1)
            year,week, _=previous_week_date.isocalendar()
            
        print(f"Report of week {week} of {year} generation...")
        create(logbook_sh,discrepancy,preventive_maintenance,rtms,week=week,year=year)
    except Exception as inputErr:
        print(inputErr)
        return False


if __name__=='__main__':
    main__() 