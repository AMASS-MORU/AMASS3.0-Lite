import pandas as pd
import gc 
import psutil, os
import math as math
from datetime import date, timedelta, datetime
from reportlab.pdfgen import canvas #creating PDF page
from reportlab.lib.units import inch #importing inch for plotting
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY #setting paragraph style
from reportlab.platypus import Table
from reportlab.platypus.paragraph import Paragraph
from reportlab.lib import colors
from reportlab.platypus.flowables import Flowable
from reportlab.lib.styles import ParagraphStyle
import matplotlib.colors as mcolors
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import AMASS_amr_const as AC
import AMASS_amr_commonlib as AL
import AMASS_amr_commonlib_report as REP_AL
import AMASS_annex_c_const as ACC

#General const
ANNEXC_RPT_CONST_TITLE = 'Annex C: Cluster signals'
ANNEXC_RPT_CONST_PVALUELIMIT = 0.05
ANXC_CONST_SP_ALL ='all'
ANXC_CONST_FOOT_REPNAME = "Supplementary Annex C"
ANXC_CONST_NUM_TOPWARD_TODISPLAY = 3
ANXC_CONST_COL_SUP_DISPLAYWARDGRAPH = "Display_graph"
ANXC_CONST_BLANK_PROFILE_TO_PREFIX = "Profile_"
ANXC_CONST_BLANK_PROFILE_TO_SUFFIX = "_00"

#Cut new page
#Annex C profile
ANXC_CONST_CLUSTERLIST_MAXROW_FIRSTPAGE = 8
ANXC_CONST_CLUSTERLIST_MAXROW_SECONDPAGEON = 27
#Supplementary profile
ANXC_CONST_PROFILE_MAXROW_FIRSTPAGE = 16
ANXC_CONST_PROFILE_MAXROW_NEXTPAGE = 22
#Supplementary ward, ward graph
ANXC_CONST_MAX_BASELINEGRAPH = 500
ANXC_CONST_MAX_BASELINEGRAPH_PERPAGE = 3
ANXC_CONST_MAX_BASELINETBLROW_PERPAGE = 27
ANXC_CONST_MAX_BASELINECOLTBL_PERPAGE =3

#Split ward/profile in baseline (SaTSCcan input)
ANXC_CONST_BASELINE_VAR_WARDPROFILE_SPLITER = ";"
ANXC_CONST_BASELINE_VAR_WARDPROFILE_SPLITID_FOR_WARD = 0
ANXC_CONST_BASELINE_VAR_WARDPROFILE_SPLITID_FOR_PROFILE = 1
#Transform specimen date in Baseline file (SaTScan input file)
ANXC_CONST_BASELINE_NEWVAR_SPECDATE = "spcdate2"
ANXC_CONST_BASELINE_NEWVAR_WEEK = "spcweek"
ANXC_CONST_COL_DAYTOBASELINEDATE = "daysbaseline"
#Transform sdate in  Cluster result file (AnnexC_listofpassedclusters_xxx_xxx.xlsx)
ANXC_CONST_COL_STARTCDATE = "startclusterdate"
ANXC_CONST_COL_ENDCDATE = "endclusterdate"
ANXC_CONST_COL_DAYTOSTARTC = "daysstartcluster"
ANXC_CONST_COL_DAYTOENDC= "daysendcluster"
#Merge baseline and cluster
ANXC_CONST_COL_PROFILE_WITHCLUSTER = "cluster_profile_id"
ANXC_CONST_PROFILENAME_FORNOCLUSTER = "Other profiles"
ANXC_CONST_NOCLUSTER_COLOR = "#CCCCCC"
ANXC_CONST_COL_GOTCLUSTER = "is_cluster"
#Grpah file name
ANXC_CONST_CLUSTER_GRAPH_FNAME = 'annexc_graph_cluster'
ANXC_CONST_BASELINE_GRAPH_FNAME = 'annexc_graph_baseline'
#New column for sum total case
ANXC_CONST_COL_PROFILESUMCASE = "total_profile_cases"
ANXC_CONST_COL_WARDSUMCASE = "total_ward_cases"
#----------------------------------------------------------------------------------------------------------
#supplementary function
class ROTATETEXT(Flowable): #TableTextRotate
    '''Rotates a tex in a table cell.'''
    def __init__(self, text ):
        Flowable.__init__(self)
        self.text=text

    def draw(self):
        canvas = self.canv
        canvas.rotate(90)
        canvas.drawString( 0, -1, self.text)

    
#----------------------------------------------------------------------------------------------------------
#Get rough total page
def get_annexC_roughtotalpage(logger):
    itotalpage = 1 + (len(ACC.dict_spc)*len(ACC.dict_org))
    try:
        icaltotalpage = 1
        for sh_org in ACC.dict_org.keys():
            #for sp in ["_all","_blo"]:
            for sp in ACC.dict_spc:
                #Cluster pages
                try:
                    df = pd.read_excel(AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_PCLUSTER+"_"+str(sh_org)+"_"+str(sp)+".xlsx")
                    x = math.ceil((len(df) - ANXC_CONST_CLUSTERLIST_MAXROW_FIRSTPAGE)/ANXC_CONST_CLUSTERLIST_MAXROW_SECONDPAGEON) + 1
                    if x <= 0:
                        x = 1
                    if x>1:
                        print("Got more than 1 page for clusters")
                    icaltotalpage  = icaltotalpage  +  x
                except:
                    AL.printlog("Warning : error calculate ANNEX C page (cluster result) for " + str(sh_org)+" : "+str(sp),False,logger)
                    icaltotalpage  = icaltotalpage  + 1
        itotalpage = icaltotalpage
    except:
        pass
    return itotalpage 
#Prepare data functions 
def annexc_table_main(df):
    #BACGROUD, (FIRSTCOL,FIRSTROW), (LASTCOL,LASTROW), color - This can set to overlap, last one will be on top of previous one
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])
def annexc_table_sup_ward(df,cl):
    #BACGROUD, (FIRSTCOL,FIRSTROW), (LASTCOL,LASTROW), color - This can set to overlap, last one will be on top of previous one
    return Table(df,colWidths=cl,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])
def annexc_table_sup_micro_ward(df,cl):
    #BACGROUD, (FIRSTCOL,FIRSTROW), (LASTCOL,LASTROW), color - This can set to overlap, last one will be on top of previous one
    return Table(df,colWidths=cl,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('ALIGN',(1,0),(-1,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])
def annexc_table_nototalrow(df):
    #BACGROUD, (FIRSTCOL,FIRSTROW), (LASTCOL,LASTROW), color - This can set to overlap, last one will be on top of previous one
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])
def annexc_table_withtotalrow(df):
    #BACGROUD, (FIRSTCOL,FIRSTROW), (LASTCOL,LASTROW), color - This can set to overlap, last one will be on top of previous one
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                           ('FONT',(0,-1),(-1,-1),'Helvetica-Bold'),
                           ('BACKGROUND',(0,-1),(-1,-1),colors.lightgrey)])
def annexc_table_nototalrow_rotate(df,rh):
    #BACGROUD, (FIRSTCOL,FIRSTROW), (LASTCOL,LASTROW), color - This can set to overlap, last one will be on top of previous one
    return Table(df,rowHeights=rh,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                           ('ALIGN',(1,0),(-2,0),'CENTER'),
                           ('VALIGN',(1,0),(-2,0),'BOTTOM')])
def annexc_table_withtotalrow_rotate(df,rh):
    #BACGROUD, (FIRSTCOL,FIRSTROW), (LASTCOL,LASTROW), color - This can set to overlap, last one will be on top of previous one
    return Table(df,rowHeights=rh,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                           ('ALIGN',(1,0),(-2,0),'CENTER'),
                           ('VALIGN',(1,0),(-2,0),'BOTTOM'),
                           ('FONT',(0,-1),(-1,-1),'Helvetica-Bold'),
                           ('BACKGROUND',(0,-1),(-1,-1),colors.lightgrey)])
                           #('ROWHEIGHT', (0, 0), (-1, -1), 1.5*inch)])

#Assign profile color done once for each org-atb
def get_colorlist():
    tab10_palette = sns.color_palette("tab10")
    list_color_names = tab10_palette.as_hex()
    return list_color_names
def get_dict_profile_color(lst_profile,list_color):
    imaxcolor = len(list_color)
    #df=pd.DataFrame(columns=["profile","color"])
    dict_pc = {}
    i = 0
    for p in lst_profile:
        scolor_name = list_color[int(i % imaxcolor)]
        #orow = {"profile":p,"color":scolor_name}
        #df=pd.concat([df,pd.DataFrame([orow])],ignore_index = True) 
        dict_pc[p] = scolor_name
        i = i + 1
    return dict_pc
#create baseline group
def get_axixlength(b_graph=False, real_max_yaxis=10):
    set_max_yaxis = 0
    step = 0
    if b_graph:
        if real_max_yaxis<=5:
            set_max_yaxis = 5
            step = 1
        elif (real_max_yaxis>5) and (real_max_yaxis<=10):
            set_max_yaxis = 10
            step = 2
        elif (real_max_yaxis>10) and (real_max_yaxis<=20):
            set_max_yaxis = 25
            step = 2
        elif (real_max_yaxis>20) and (real_max_yaxis<=30):
            set_max_yaxis = 40
            step = 5
        elif (real_max_yaxis>30) and (real_max_yaxis<=50):
            set_max_yaxis = 60
            step = 5
        else:
            set_max_yaxis = real_max_yaxis+30
            step = 5
    else:
        pass
    return set_max_yaxis, step
def sub_printprocmem(sstate,logger) :
    try:
        process = psutil.Process(os.getpid())
        AL.printlog("Memory usage at state " +sstate + " is " + str(process.memory_info().rss) + " bytes.",False,logger) 
    except:
        AL.printlog("Error get process memory usage at " + sstate,True,logger)
def create_graph_baseline(raw_df, swardid,scolweek,scolnumcase,sfilename,logger):
    sub_printprocmem("Baseline graph ",logger)
    plt.figure(figsize=(15,3))
    axislength = get_axixlength(b_graph=len(raw_df.columns)>1, real_max_yaxis=raw_df[scolnumcase].max())
    #raw_df.plot(x=scolweek,y=scolnumcase,kind='bar', fontsize=10, yticks=np.arange(0,axislength[0]+1,step=axislength[1]),legend=False,color=['gray'])
    ax=sns.barplot(x=scolweek,y=scolnumcase,data=raw_df,color='#E76F51')
    ax.set_yticks(np.arange(0,axislength[0]+1,step=axislength[1]))
    ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
    plt.ylabel("Number of patients(n)*", fontsize=10)
    plt.xlabel("Number of week (Start of weekday)", fontsize=10)
    plt.savefig(sfilename, format='png',dpi=200,transparent=True, bbox_inches="tight")
    plt.close()
    plt.clf
#Main function for prepare data
def main_preparedata(logger):
    return True

#----------------------------------------------------------------------------------------------------------
# Generate ANNEX C report in PDF
def template_page1_intro():
    return ["This supplementary report shows the information of potential clusters which are identified using the SatScan (www.satscan.org). " + 
            "An outbreak of hospital-acquired infection (HAI) may be defined as an increase in the occurrence of HAI compared to a recorded baseline rate.", 
            "<br/>This report is generated by default, even if hospital_admission_data is unavailable. This is to enable hospitals with only microbiology data to utilize the de-duplication and automation of AMASS-SatScan and report generation functions of AMASS.",
            "<br/>The AMASS-SatScan uses an automation and tier-based approach. In case when the hospital admission data file is available, the AMASS-SatScan will include only hospital-origin infections in the analysis. In case when the hospital admission data file is not available, the AMASS-SatScan will include all infections (without stratification of infection origin) in the analysis.  In case when the microbiology data file has ward data, the AMASS-SatScan will identify cluster signals using ward data and consider each ward as an independent ward. In case when the microbiology data file does not have the ward data, the AMASS-SatScan will identify cluster signals by considering that the hospital as a single space. ",
            "<br/>Please note that the AMASS-SatScan de-duplicated the data by including only the first resistant isolate per patient per specimen type per evaluation period. Therefore, the total number of patients with resistant infections in this section may be minimally higher than those reported in the section two to six.",
            "<br/>Please note that cluster signals may be false signals and that actual outbreaks may not be identified due to many factors such as data anomalies, data errors, limitation of diagnostic practice and confounding factors. Please refer to the ‘Methods’ section for more details on the default setting for the AMASS-SatScan."]
def template_page1_org():
    lst_org = []
    for skey in ACC.dict_org:
        lorg = ACC.dict_org[skey]
        sorg = lorg[1] + " (" + lorg[2] + ")"
        lst_org = lst_org + [sorg]
    return lst_org
def OBSOLETED_template_page1_org():
    return ["Methicillin-R <i>S. aureus</i> (MRSA)",
            "Vancomycin-R <i>Enterococcus</i> spp. (VRE) for <i>E. faecalis</i> and <i>E. faecium</i>",
            "Carbapenem-R Enterobacteriaceae (CRE) for <i>E. coli</i> and <i>K. pneumoniae</i>",
            "Carbapenem-R <i>P. aeruginosa</i> (CRPA)",
            "Carbapenem-R <i>A. baumannii</i> (CRAB)"]
def template_page1_result():
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    return [""]
def template_page2_header():
    return [["<b>" + "annexc_graph_1 : "       + "annexc_organism" + "</b>"],
            ["<b>" + "No. of patients with "   + "annexc_organism" + " : " + "${blo_num}$"  + "</b>"],
            ["<b>" + "No. of wards having patients with " + "annexc_organism" + " : " + "${blo_ward}$" + "</b>"],
            ["<b>" + "Model : "             "${blo_model}$"   + "</b>"]]
def template_page2_footer():
    return ["<font color=darkgreen>" + "annexc_footer_1" + " "  + 
            "annexc_footer_2"  + " "  + "annexc_footer_3"  + " "  + "annexc_footer_4"  + " " +
            "annexc_footer_ci" + "; " + "annexc_footer_ns" + "; " + "annexc_footer_na" + " " +
            "annexc_footer_ast" + "</font>"]
def footnote_annexC(sp):
    dict_note = {"blo":
                 ['* The AMASS-SatScan (Annex C) de-duplicated by including only the first resistant isolate per patient per specimen type per evaluation period. ' +
                         'Bar graphs show patients with blood culture positive with organism profiles which were identified in at least one cluster signal. ' +
                         'Gray bars (Other profiles) represents patients with blood culture positive for organisms with profiles that were not included in any cluster signals. ' + 
                         'Details of AMR profiles are available in "Supplementary_data_Annex_C.pdf" and files in the folder "Report_with_patient_identifiers"']
                }
    if sp in dict_note:
        return dict_note[sp]
    else:
        return ['* The AMASS-SatScan (Annex C) de-duplicated by including only the first resistant isolate per patient per specimen type per evaluation period. ' +
                        'Bar graphs show patients with a clinical specimen  culture positive with organism profiles which were identified in at least one cluster signal. ' +
                        'Gray bars (Other profiles) represents patients with a clinical specimen positive for organisms with profiles that were not included in any cluster signals. ' + 
                        'Details of AMR profiles are available in "Supplementary_data_Annex_C.pdf" and files in the folder "Report_with_patient_identifiers"']
def footnote_baseline(sp):
    return footnote_annexC(sp)


# Function to replace non-integers with 0
def replace_non_integer(value):
    try:
        # Attempt to convert the value to an integer
        int_value = int(value)
        return int_value
    except (ValueError, TypeError):
        # If it's not an integer or can't be converted, replace with 0
        return 0
def prapare_mainAnnexC_mainpage(canvas_rpt,logger,page,startpage,lastpage,totalpage,strgendate):
    #c = canvas.Canvas(filename_pdf+"_"+str(1)+".pdf")
    REP_AL.report_title(canvas_rpt,ANNEXC_RPT_CONST_TITLE,1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    REP_AL.report_title(canvas_rpt,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    REP_AL.report_context(canvas_rpt, template_page1_intro(),                                           
                   1.0*inch, 6.6*inch, 460, 200, font_size=11, font_align=TA_LEFT)
    REP_AL.report_title(canvas_rpt,'Pathogens under the survey',1.07*inch, 6.2*inch,'#3e4444',font_size=12)
    REP_AL.report_context(canvas_rpt, template_page1_org(),                                           
                   1.0*inch, 4.0*inch, 460, 150, font_size=11, font_align=TA_LEFT)
    REP_AL.report_title(canvas_rpt,'Results',1.07*inch, 3.1*inch,'#3e4444',font_size=12)
    REP_AL.report_context(canvas_rpt, template_page1_result(),                                           
                   1.0*inch, 1.2*inch, 460, 200, font_size=11, font_align=TA_LEFT)
    #REP_AL.report_todaypage(canvas_rpt,270,30,"AnnexC - Page " + str(1))
    REP_AL.canvas_printpage(canvas_rpt,page,lastpage,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,'Annex C',totalpage,startpage)
    #canvas_rpt.showPage()
    #c.save()
def save_generalinfo(df,sh_org,spec,p,v):
    dict_tosave = {'Pathogen':sh_org,'Specimen':spec,'Parameters':str(p),'Values':str(v)}
    new_row=pd.DataFrame([dict_tosave])
    df = pd.concat([df, new_row], ignore_index=True)
    return df
    

def getyearweek(ddate):
    return str(ddate.strftime("%Y")) + "_" + str(ddate.strftime("%W"))
def create_df_weekday(s_studydate="2021/01/01", e_studydate="2021/12/31", fmt_studydate="%Y/%m/%d",
                              col_sweekday="startweekday",col_yearweek="year_week"):
    
    ds = datetime.strptime(s_studydate,fmt_studydate)
    ds_monday = (ds - timedelta(days=ds.weekday()))
    de = datetime.strptime(e_studydate,fmt_studydate)
    de_monday = (de - timedelta(days=de.weekday()))
    numweek = math.ceil((de_monday - ds_monday).days / 7)
    df = pd.DataFrame("",
                        index=range(0,numweek), 
                        columns=[col_yearweek,col_sweekday])
    for idx in df.index:
        dtmp_monday = ds_monday + timedelta(days=7*idx)
        df.at[idx,col_sweekday] = str(idx) +" (" + dtmp_monday.strftime("%Y-%m-%d") + ")" 
        df.at[idx,col_yearweek] = getyearweek(dtmp_monday) 
    return df
def caldays(df,coldate,dbaselinedate) :
    return (df[coldate] - dbaselinedate).dt.days
def prepare_tabletoplot(logger,df,lst_col=[],lst_tosort=[], list_ascending=[],dict_displaycol={},col_checkgroup="",col_tohidevalue=""):
    try:
        #df = df.merge(df_profile_sum, how="left", left_on=ACC.CONST_COL_PROFILEID, right_on=ACC.CONST_COL_PROFILEID,suffixes=("","_SUM"))
        #df[ANXC_CONST_COL_PROFILESUMCASE] = df[ANXC_CONST_COL_PROFILESUMCASE].fillna(0)
        #print(df)
        df = df.sort_values(by = lst_tosort, ascending = list_ascending, na_position = "last")
        df = df[lst_col]
        if (col_checkgroup!="") & (col_tohidevalue!=""):
            if (col_checkgroup in lst_col) & (col_tohidevalue in lst_col):
                #Don't display total per ward if same ward
                mask = df[col_checkgroup] == df[col_checkgroup].shift()
                df.loc[mask, col_tohidevalue] = ''
        if len(dict_displaycol) > 0:
            df.rename(columns=dict_displaycol, inplace=True)
    except Exception as e:                #has no cluster >>> ignore
        AL.printlog("Error : checkpoint Annex C generate report (prepare cluster table) : " +str(e),True,logger) 
        logger.exception(e)
    return df
def gen_cluster_graph(logger,df=pd.DataFrame(),xcol_sort="",xcol_display="",ycol="",zcol="",list_privotcol=[], dict_profile_color={}, filename="",
         xlabel="Number of week (Start of weekday)",ylabel="Number of patients(n)*",figsizex=20,figsizey=10):
    #df_forsave=pd.DataFrame()
    df_sum = df.groupby([xcol_display])[ACC.CONST_COL_CASE].sum().reset_index(name=ACC.CONST_COL_CASE)
    axislength = get_axixlength(b_graph=len(df.columns)>1, real_max_yaxis=df_sum[ACC.CONST_COL_CASE].max())
    plt.figure()
    xcol = xcol_display
    if xcol_sort.strip() != "":
        xcol = [xcol_display,xcol_sort]
    else:
        xcol_sort = xcol_display
    df_org = df.pivot_table(index=xcol, columns=zcol, values=ycol, aggfunc='sum', fill_value=0)
    temp_list = [x for x in list_privotcol if x in df_org.columns.to_list()]
    temp_list2 = [x for x in df_org.columns.to_list() if x not in temp_list]
    col_list = temp_list + temp_list2
    df_org = df_org.sort_values(by=[xcol_sort], ascending=[True])
    df_org.reset_index(inplace=True)
    df_org.set_index(xcol_display, inplace=True)
    df_org = df_org[col_list]
    palette = [dict_profile_color.get(item, item) for item in df_org.columns.tolist()]
    df_org.plot(kind='bar', 
            stacked =True, 
            figsize =(figsizex, figsizey), 
            color =palette, 
            fontsize=16)
    plt.legend(prop={'size': 14},ncol=4)
    plt.ylabel(ylabel, fontsize=16)
    plt.xlabel(xlabel, fontsize=16)
    plt.yticks(np.arange(0,axislength[0]+1,step=axislength[1]), fontsize=16)
    plt.savefig(filename, format='png',dpi=180,transparent=True, bbox_inches="tight")
    plt.close()
    plt.clf
    #sub_printprocmem(filename,logger)
    #return(df_org)
#-----------------------------------------------------------------------------------------------------------
#Generate part of the report
def cover(c,logger,strgendate):
    sec1_res_i = AC.CONST_FILENAME_sec1_res_i
    sec1_file = AC.CONST_PATH_RESULT + sec1_res_i
    section1_result = pd.DataFrame()
    try:
        section1_result = pd.read_csv(sec1_file).fillna("NA")
    except Exception as e:
        AL.printlog("Error read file : " + sec1_file + " : " + str(e),True,logger)
        return 
    if len(section1_result) <= 0:
        AL.printlog("Error no record in file : " + sec1_file ,True,logger)
        return
    if len(section1_result) > 0:
        hospital_name   = REP_AL.assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Hospital_name","Values"].tolist()[0]), coverpage=True)
        country_name    = REP_AL.assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Country","Values"].tolist()[0]), coverpage=True)
        spc_date_start  = REP_AL.assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="overall_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0]), coverpage=True)
        spc_date_end    = REP_AL.assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="overall_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0]), coverpage=True)
    else:
        hospital_name    = "NA"
        country_name    = "NA"
        spc_date_start  = "NA"
        spc_date_end    = "NA"
    ##paragraph variable
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    add_blankline = "<br/>"
    ##content
    cover_1_1 = "<b>Hospital name:</b>  " + bold_blue_op + hospital_name + bold_blue_ed
    cover_1_2 = "<b>Country name:</b>  " + bold_blue_op + country_name + bold_blue_ed
    cover_1_3 = "<b>Data from:</b>"
    cover_1_4 = bold_blue_op + str(spc_date_start) + " to " + str(spc_date_end) + bold_blue_ed
    cover_1 = [cover_1_1,cover_1_2,add_blankline+cover_1_3, cover_1_4]
    cover_2_1 = "This is a detailed report for records with cluster signals identified by the AMASS-SatScan. This report, together with the full list in Excel format, is for users to check and validate the cluster and the patients in each cluster identified by the SatScan. The information available in this PDF file include ward names used in the dictionary files. The identifiers in the Excel files for the Annex C include hospital number and specimen collection date. Users should not share or transfer this report and the excel files for the Annex C to any party outside of the hospital without data security management and confidential agreement."
    cover_2_2 = "<br/><b>Generated on:</b>  " + bold_blue_op + strgendate + bold_blue_ed
    cover_2 = [cover_2_1,cover_2_2]
    ##reportlab
    c.setFillColor('#FCBB42')
    c.rect(0,590,800,20, fill=True, stroke=False)
    c.setFillColor(colors.royalblue)
    c.rect(0,420,800,150, fill=True, stroke=False)
    REP_AL.report_title(c,'Supplementary report Annex C:',0.7*inch, 515,'white',font_size=28)
    REP_AL.report_title(c,'Cluster signals identify by AMASS-SatScan',0.7*inch, 455,'white',font_size=20)
    REP_AL.report_context(c,cover_1, 0.7*inch, 3.0*inch, 460, 180, font_size=18,line_space=26)
    REP_AL.report_context(c,cover_2, 0.7*inch, 0.5*inch, 460, 120, font_size=10,line_space=13)
    c.showPage()
def prapare_supplementAnnexC_per_org(canvas_sup_rpt,logger,page,startpage,lastpage,totalpage,strgendate,sh_org="",spec="",df=pd.DataFrame(),df_baseline=pd.DataFrame(),df_profile=pd.DataFrame(),df_ward=pd.DataFrame(),list_profile_atb_column=[]):
    #noclustertext = "There is no cluster with p-value < " +str(pvaluelimit) + "."
    #shaveclustertext = "Table of clusters with p-value < " +str(pvaluelimit) + "."
    style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_CENTER)
    sspecname = ACC.dict_spc[spec]
    sorgname = ACC.dict_org[sh_org][2]
    dict_dis_specname= {"blo":"blood",
                        "all":"clinical specimens"}
    sdis_spec = dict_dis_specname[spec]
    sheader_wardgrap = f'Display of patients with {sdis_spec} culture positive for {sorgname} in each ward over time'
    sfootnote = "* The AMASS-SatScan (Annex C) de-duplicated by including only the first resistant isolate per patient per specimen type per evaluation period. "
    sfootnote = sfootnote + f"Bar graphs show patients with {sdis_spec} culture positive with the organism with a profile identified in at least one cluster signal. "
    sfootnote = sfootnote + "Gray bars (Other profiles) represents patients with blood culture positive for organisms with profiles that were not included in any cluster signals. "
    sfootnote = sfootnote + "Only wards with a cluster signal identified or having the top three highest number of patients were displayed. "
    lst_footnote = [sfootnote]
    lst_footnote_wardlist = ["* In case that there are ward names in your hospital_admission_data file, this list and the analysis will prioritize the ward names in the microbiology_data file over the ones in hospital_admission_data file. "]
    
    stotal_patient = "NA"
    stotal_ward = "NA"
    stotal_profile = "NA"
    scluster_patient = "NA"
    scluster_ward = "NA"
    scluster_profile = "NA"
    try:
        stotal_patient = str(df_baseline[ACC.CONST_COL_CASE].sum())
        df_ward_sum = df_baseline.groupby(by=ACC.CONST_COL_WARDID)[ACC.CONST_COL_CASE].sum().reset_index(name=ACC.CONST_COL_CASE)
        df_ward_sum.rename(columns={ACC.CONST_COL_CASE:ANXC_CONST_COL_WARDSUMCASE}, inplace=True)
        stotal_ward = str(len(df_ward_sum))
        #stotal_ward = str(len(df_baseline.groupby(by=ACC.CONST_COL_WARDID)[ACC.CONST_COL_WARDID].count()))
        stotal_profile = str(len(df_baseline.groupby(by=ACC.CONST_COL_PROFILEID)[ACC.CONST_COL_PROFILEID].count()))
    except Exception as e:
        AL.printlog("Warning : Baseline data not available for total sum : " + str(spec) + " of " + str(sh_org) ,False,logger) 
    #Keep may be useful if add more info to baseline like in Annex C
    REP_AL.report_title(canvas_sup_rpt,sspecname +": " + sorgname,1.07*inch, 10.6*inch,'#3e4444',font_size=16)
    #General info ---------------------------------------
    REP_AL.report_context(canvas_sup_rpt, ["<b>"+ "Baseline information" + "</b>"],                                       
                   1.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
    REP_AL.report_context(canvas_sup_rpt, ["<i>"+ f'No. of patients = {stotal_patient}' + "</i>"],
                   1.5*inch, 9.2*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    REP_AL.report_context(canvas_sup_rpt, ["<i>"+ f'No. of wards = {stotal_ward}' + "</i>"],
                   1.5*inch, 9.0*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    REP_AL.report_context(canvas_sup_rpt, ["<i>"+ f'No. of AMR profiles = {stotal_profile}' + "</i>"],
                   1.5*inch, 8.8*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    #--------------------------------------------------------------------------------------------------------------
    #Profile
    try:
        if len(df_profile) > 0:
            df_profile.sort_values(by = [ANXC_CONST_COL_PROFILESUMCASE], ascending = [False], na_position = "last")
            try:
                dict_total= {}
                for scol in list_profile_atb_column:
                    dict_total[scol] ="-"
                    #The following code will be useful for cal sum of each atb in the future don't delete
                    """
                    dict_total[scol] =0
                    try:
                        scol_count = scol+"_count"
                        df_profile[scol_count] =  df_profile[ANXC_CONST_COL_PROFILESUMCASE]
                        df_profile.loc[(df_profile[scol] == '-') | (df_profile[scol] == 'ND'), scol_count] = 0
                        dict_total[scol] = df_profile[scol_count].sum()
                    except:
                        pass
                    try:
                        df_profile.drop(scol_count, axis=1, inplace=True)
                    except:
                        pass
                    """
                dict_total[ANXC_CONST_COL_PROFILESUMCASE] = df_profile[ANXC_CONST_COL_PROFILESUMCASE].sum()
                dict_total[ANXC_CONST_COL_PROFILESUMCASE] = Paragraph("<b>" + str(df_profile[ANXC_CONST_COL_PROFILESUMCASE].sum()) + "</b>",style_summary)
                dict_total[ACC.CONST_COL_PROFILEID] = Paragraph("<b>" + "Total"+ "</b>",style_summary)
                new_total_row = pd.DataFrame([dict_total])
                df_profile = pd.concat([df_profile, new_total_row], ignore_index=True)
            except Exception as e:
                AL.printlog("Warning : unable to generate total sum: " + str(spec) + " of " + str(sh_org) ,False,logger) 
            #rename column
            try:
                df_profile.rename(columns={ACC.CONST_COL_PROFILEID:"Profile ID",ANXC_CONST_COL_PROFILESUMCASE:"No. of\npatients"}, inplace=True)
            except Exception as e:
                pass
            #prepare rotate column header
            lst_column_head = []
            atbname_c = 0
            for scol in df_profile.columns.tolist():
                if len(scol) > atbname_c:
                    atbname_c = len(scol)
                if scol in list_profile_atb_column:
                    lst_column_head = lst_column_head + [ROTATETEXT(scol)]
                else:
                    lst_column_head = lst_column_head + [scol]
            #print table
            iheader_h = 0.068*atbname_c
            #Print report
            REP_AL.report_context(canvas_sup_rpt, ["<b>"+ "List of profiles" + "</b>"],                                       
                           1.0*inch, 8.2*inch, 460, 50, font_size=13, font_align=TA_LEFT)
            df1 = df_profile[:ANXC_CONST_PROFILE_MAXROW_FIRSTPAGE]
            ioffset = 0.25*(len(df1)) + iheader_h
            lst_df = [lst_column_head] + df1.values.tolist()
            rh = [iheader_h*inch]
            for i in range(len(df1)):
                rh = rh + [0.25*inch]
            table_draw = annexc_table_nototalrow_rotate(lst_df,rh)
            table_draw.wrapOn(canvas_sup_rpt, 480+ioffset, 300)
            table_draw.drawOn(canvas_sup_rpt, 1.1*inch, (8.2-ioffset)*inch)
            REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
            REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
            df_profile = df_profile[ANXC_CONST_PROFILE_MAXROW_FIRSTPAGE:]
            page = page + 1
            if len(df_profile)>0:
                imorepage = math.ceil(len(df_profile)/ANXC_CONST_PROFILE_MAXROW_NEXTPAGE)
                for i in range(imorepage):
                    REP_AL.report_title(canvas_sup_rpt,sspecname +": " + sorgname,1.07*inch, 10.6*inch,'#3e4444',font_size=16)
                    REP_AL.report_context(canvas_sup_rpt, ["<b>"+ "List of profiles (Continue)" + "</b>"],                                       
                                   1.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
                    df1 = df_profile[:ANXC_CONST_PROFILE_MAXROW_NEXTPAGE]
                    ioffset = 0.25*(len(df1)) + iheader_h
                    lst_df = [lst_column_head] + df1.values.tolist()
                    rh = [iheader_h*inch]
                    for i in range(len(df1)):
                        rh = rh + [0.25*inch]
                    table_draw = annexc_table_nototalrow_rotate(lst_df,rh)
                    #table_draw._awgH[0] = iheader_h*inch
                    table_draw.wrapOn(canvas_sup_rpt, 480+ioffset, 300)
                    table_draw.drawOn(canvas_sup_rpt, 1.1*inch, (9.5-ioffset)*inch)
                    REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
                    REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
                    df_profile = df_profile[ANXC_CONST_PROFILE_MAXROW_NEXTPAGE:]
                    page = page + 1
        else:
            REP_AL.report_context(canvas_sup_rpt, ["<b>"+ "list of profiles" + "</b>"], 
                   1.0*inch, 8.2*inch, 460, 50, font_size=11, font_align=TA_LEFT)
            REP_AL.report_context(canvas_sup_rpt, ["None"],
                           1.5*inch, 7.9*inch, 460, 50, font_size=11, font_align=TA_LEFT)
            REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
            REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
            page = page + 1
    except:
        REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
        REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
        page = page + 1
    #--------------------------------------------------------------------------------------------------------------
    #ward
    #Generate report page (Loop for table)
    df = df_ward.copy(deep=True)
    if len(df) > 0:
        try:
            dict_total = {}
            dict_total[ANXC_CONST_COL_WARDSUMCASE] = Paragraph("<b>" + str(df[ANXC_CONST_COL_WARDSUMCASE].sum()) + "</b>",style_summary)
            dict_total[ACC.CONST_COL_WARDID] = Paragraph("<b>" + "Total"+ "</b>",style_summary)
            new_total_row = pd.DataFrame([dict_total])
            df = pd.concat([df, new_total_row], ignore_index=True)
        except Exception as e:
            AL.printlog("Warning : checkpoint Annex C generate report (ward summary) " + str(spec) + ") of " + str(sh_org) + " : " +str(e),False,logger) 
            logger.exception(e)
        try:
            df = df[[ACC.CONST_COL_WARDID,ANXC_CONST_COL_WARDSUMCASE]]
            df.rename(columns={ACC.CONST_COL_WARDID:"Ward ID",ANXC_CONST_COL_WARDSUMCASE:"No. of\npatients"}, inplace=True)
        except Exception as e:
            pass
        icurcol = 1
        icoloffset = 2.5
        balreadyprintpage = False
        bisfirstpage = True
        while len(df) > 0:
            if icurcol == 1:
                #print header
                balreadyprintpage = False
                REP_AL.report_title(canvas_sup_rpt,sspecname +": " + sorgname,1.07*inch, 10.6*inch,'#3e4444',font_size=16)
                scon = " (Continue)" if bisfirstpage == False else ""
                REP_AL.report_context(canvas_sup_rpt, ["<b>"+ "List of ward" + scon + "</b>"],                                       
                               1.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
            df1 = df[:ANXC_CONST_MAX_BASELINETBLROW_PERPAGE]
            ioffset = 0.25*(len(df1) + 1)
            try:
                lst_df = [df1.columns.tolist()] + df1.values.tolist()
                table_draw = annexc_table_sup_ward(lst_df,[1*inch,1*inch])
                table_draw.wrapOn(canvas_sup_rpt, 480+ioffset, 300)
                table_draw.drawOn(canvas_sup_rpt, (1+((icurcol-1)*icoloffset))*inch, (9.3-ioffset)*inch)
            except Exception as e:
                AL.printlog("Error : checkpoint Annex C generate report (table) " + str(spec) + ") of " + str(sh_org) + " : " +str(e),True,logger) 
                logger.exception(e)
            df = df[ANXC_CONST_MAX_BASELINETBLROW_PERPAGE:]
            if icurcol >= ANXC_CONST_MAX_BASELINECOLTBL_PERPAGE-1:
                if balreadyprintpage == False:
                    REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
                    REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
                    bisfirstpage = False
                    page = page + 1
                    icurcol = 1
                    balreadyprintpage = True   
            else:
                icurcol = icurcol + 1
        if balreadyprintpage == False:
            REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
            REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
            bisfirstpage = False
            page = page + 1
            icurcol = 1
            balreadyprintpage = True
    else:
        REP_AL.report_title(canvas_sup_rpt,sspecname +": " + sorgname,1.07*inch, 10.6*inch,'#3e4444',font_size=16)
        REP_AL.report_context(canvas_sup_rpt, ["<b>"+ "list of wards" + "</b>"],                                       
                       1.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
        REP_AL.report_context(canvas_sup_rpt, ["None"],
                       1.5*inch, 9.2*inch, 460, 50, font_size=11, font_align=TA_LEFT)
        REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
        REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
        page = page + 1
    #--------------------------------------------------------------------------------------------------------------
    #ward graph
    #generate report page
    if len(df_ward) > 0:
        ifigoffset_inch = 2.50
        ifirstpos_inch = 9.3
        igc = 0
        icurpos = 1
        balreadyprintpage = False
        #print(df_table)
        #generate report page (Loop for graphs)
        for idx in df_ward.index:
            #if igc < ANXC_CONST_MAX_BASELINEGRAPH:
            #df_ward.loc[idx,ACC.CONST_COL_WARDID]
            if df_ward.loc[idx,ANXC_CONST_COL_SUP_DISPLAYWARDGRAPH]> 0:
                swardid = str(df_ward.loc[idx,ACC.CONST_COL_WARDID])
                icurpos = (igc % ANXC_CONST_MAX_BASELINEGRAPH_PERPAGE) + 1
                if icurpos == 1:
                    balreadyprintpage = False
                    REP_AL.report_title(canvas_sup_rpt,sspecname +": " + sorgname,1.07*inch, 10.6*inch,'#3e4444',font_size=16)
                    REP_AL.report_context(canvas_sup_rpt, ["<b>"+ sheader_wardgrap +"</b>"],                                           
                                   1.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
                REP_AL.report_context(canvas_sup_rpt, ["<b><i><font color=\"#000080\">"+"Ward : " + swardid +"</font></i></b>"],              
                               1.0*inch, (ifirstpos_inch-((icurpos-1)*ifigoffset_inch)-0.3)*inch, 460, 50, font_size=11, font_align=TA_LEFT)
                try:
                    canvas_sup_rpt.drawImage(AC.CONST_PATH_TEMPWITH_PID + ANXC_CONST_BASELINE_GRAPH_FNAME + "_" + str(sh_org) + "_" + str(spec) + "_" + str(swardid) + ".png", 
                                         (0-1)*inch, (ifirstpos_inch-(icurpos*ifigoffset_inch)+0.4) *inch, preserveAspectRatio=True, width=10*inch, height=2*inch,showBoundary=False) 
                except Exception as e:
                    AL.printlog("Error : checkpoint Annex C generate report (graph) " + str(spec) + ") of " + str(sh_org) + " : " +str(e),True,logger) 
                    logger.exception(e)
                if (icurpos >= ANXC_CONST_MAX_BASELINEGRAPH_PERPAGE):
                    REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
                    REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
                    page = page + 1
                    balreadyprintpage = True
                igc = igc+1
        if balreadyprintpage == False:
            REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
            REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
            page = page + 1
            balreadyprintpage = True
    else:
        REP_AL.report_title(canvas_sup_rpt,sspecname +": " + sorgname,1.07*inch, 10.6*inch,'#3e4444',font_size=16)
        REP_AL.report_context(canvas_sup_rpt, ["<b>"+ sheader_wardgrap +"</b>"],                                      
                       1.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
        REP_AL.report_context(canvas_sup_rpt, ["None"],
                       1.5*inch, 9.0*inch, 460, 50, font_size=11, font_align=TA_LEFT)
        REP_AL.report_context(canvas_sup_rpt,lst_footnote, 1.0*inch, 0.30*inch, 460, 130, font_size=9,line_space=12)
        REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
        page = page + 1        
    return page
def prapare_mainAnnexC_per_org(canvas_rpt,logger,page,startpage,lastpage,totalpage,strgendate,pvaluelimit=0.05,sh_org="",spec="",df=pd.DataFrame(),df_baseline=pd.DataFrame(),df_profile_sum=pd.DataFrame(), filename_graph=""):
    snoclustertext = "There is no cluster signal with p-value < " +str(pvaluelimit) + "."
    shaveclustertext = "Wards with cluster signals with p value < " +str(pvaluelimit) + "."
    sspecname = ACC.dict_spc[spec]
    sorgname = ACC.dict_org[sh_org][2]
    stotal_patient = "NA"
    stotal_ward = "NA"
    stotal_profile = "NA"
    scluster_patient = "NA"
    scluster_ward = "NA"
    scluster_profile = "NA"
    spercent_patient = "NA"
    spercent_ward = "NA"
    spercent_profile = "NA"
    df_ward_sum = pd.DataFrame(columns=[ACC.CONST_COL_WARDID,ANXC_CONST_COL_WARDSUMCASE])
    df_ward_cluster_sum = pd.DataFrame(columns=[ACC.CONST_COL_WARDID,ANXC_CONST_COL_WARDSUMCASE])
    df_resultdata_forsave = pd.DataFrame()
    try:
        stotal_patient = str(df_baseline[ACC.CONST_COL_CASE].sum())
        df_ward_sum = df_baseline.groupby(by=ACC.CONST_COL_WARDID)[ACC.CONST_COL_CASE].sum().reset_index(name=ACC.CONST_COL_CASE)
        df_ward_sum.rename(columns={ACC.CONST_COL_CASE:ANXC_CONST_COL_WARDSUMCASE}, inplace=True)
        stotal_ward = str(len(df_ward_sum))
        #stotal_ward = str(len(df_baseline.groupby(by=ACC.CONST_COL_WARDID)[ACC.CONST_COL_WARDID].count()))
        stotal_profile = str(len(df_baseline.groupby(by=ACC.CONST_COL_PROFILEID)[ACC.CONST_COL_PROFILEID].count()))
    except Exception as e:
        AL.printlog("Warning : Baseline data not available for total sum : " + str(spec) + " of " + str(sh_org) ,False,logger) 
    try:
        scluster_patient = str(df[ACC.CONST_COL_NEWOBS].sum())
        scluster_ward = str(len(df.groupby(by=ACC.CONST_COL_WARDID)[ACC.CONST_COL_WARDID].count()))
        scluster_profile = str(len(df.groupby(by=ACC.CONST_COL_PROFILEID)[ACC.CONST_COL_PROFILEID].count()))
        df_ward_cluster_sum = df.groupby(by=ACC.CONST_COL_WARDID)[ACC.CONST_COL_NEWOBS].sum().reset_index(name=ACC.CONST_COL_NEWOBS)
        df_ward_cluster_sum.rename(columns={ACC.CONST_COL_NEWOBS:ANXC_CONST_COL_WARDSUMCASE}, inplace=True)
    except Exception as e:
       AL.printlog("Warning : Cluster signal data not available for total sum : " + str(spec) + " of " + str(sh_org) ,False,logger) 
    try:
        spercent_patient = 0 if int(stotal_patient) == 0 else math.ceil(100*int(scluster_patient)/int(stotal_patient))
    except Exception as e:
        AL.printlog("Warning : either baseline or cluster signal data not available for calculate percentage (Patient) : " + str(spec) + " of " + str(sh_org) ,False,logger) 
    try:
        spercent_ward = 0 if int(stotal_ward) == 0 else math.ceil(100*int(scluster_ward)/int(stotal_ward))
    except Exception as e:
        AL.printlog("Warning : either baseline or cluster signal data not available for calculate percentage (Ward) : " + str(spec) + " of " + str(sh_org),False,logger) 
    try:
        spercent_profile = 0 if int(stotal_profile) == 0 else math.ceil(100*int(scluster_profile)/int(stotal_profile))
    except Exception as e:
        AL.printlog("Warning : either baseline or cluster signal data not available for calculate percentage (Profile) : " + str(spec) + " of " + str(sh_org),False,logger) 
    
    #get data for save in resultdata folder
    df_general_info_forsave = pd.DataFrame(columns=['Pathogen','Specimen','Parameters','Values'])
    df_general_info_forsave = save_generalinfo(df_general_info_forsave,sh_org,spec,"Total number of patients",stotal_patient)
    df_general_info_forsave = save_generalinfo(df_general_info_forsave,sh_org,spec,"Total number of patients within clusters signal identified by SaTScan method",scluster_patient)
    df_general_info_forsave = save_generalinfo(df_general_info_forsave,sh_org,spec,"Total number of profiles",stotal_profile)
    df_general_info_forsave = save_generalinfo(df_general_info_forsave,sh_org,spec,"Total number of profiles within clusters signal identified by SaTScan method",scluster_profile)
    df_general_info_forsave = save_generalinfo(df_general_info_forsave,sh_org,spec,"Total number of wards",stotal_ward)
    df_general_info_forsave = save_generalinfo(df_general_info_forsave,sh_org,spec,"Total number of wards within clusters signal identified by SaTScan method",scluster_ward)
    #---------
    REP_AL.report_title(canvas_rpt,ANNEXC_RPT_CONST_TITLE,1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    REP_AL.report_context(canvas_rpt, ["<b>"+ sspecname + ": " + sorgname + "</b>"],                                       
                   1.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
    REP_AL.report_context(canvas_rpt, ["<b>"+ "Hospital-origin" + "</b>"],                                       
                   6.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
    #General info ---------------------------------------
    REP_AL.report_context(canvas_rpt, ["<i>"+ f'No. of patients = {stotal_patient} ({scluster_patient} [{spercent_patient}%] were included in cluster signals)' + "</i>"],
                   1.5*inch, 9.2*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    REP_AL.report_context(canvas_rpt, ["<i>"+ f'No. of wards = {stotal_ward} ({scluster_ward} [{spercent_ward}%] were included in cluster signals)' + "</i>"],
                   1.5*inch, 9.0*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    REP_AL.report_context(canvas_rpt, ["<i>"+ f'No. of AMR profiles = {stotal_profile} ({scluster_profile} [{spercent_profile}%] were included in cluster signals)' + "</i>"],
                   1.5*inch, 8.8*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    
    #-----------------------------------------------------
    if os.path.exists(filename_graph):
        try:
            canvas_rpt.drawImage(filename_graph, (0-2.6)*inch, 4.80*inch, preserveAspectRatio=True, width=13.5*inch, height=4.15*inch,showBoundary=False) 
        except Exception as e:
            AL.printlog("Warning : Unable to draw graoph: " +  filename_graph + " for " + str(spec) + " of " + str(sh_org),False,logger) 
            #AL.printlog("Error : checkpoint Annex C generate report (cluster graph) " + str(spec) + ") of " + str(sh_org) + " : " +str(e),True,logger) 
            logger.exception(e)
            #print (e)
    else:
        AL.printlog("Warning : Missing generated graoph: " +  filename_graph + " for " + str(spec) + " of " + str(sh_org),False,logger) 
    #page = prepare_baselinesummary_per_org(canvas_rpt,logger,page,startpage,lastpage,totalpage,strgendate,sh_org,sp,df_week)
    if len(df)>0:
        #print(len(df))
        try:
            df = df.merge(df_profile_sum, how="left", left_on=ACC.CONST_COL_PROFILEID, right_on=ACC.CONST_COL_PROFILEID,suffixes=("","_SUM"))
            df[ANXC_CONST_COL_PROFILESUMCASE] = df[ANXC_CONST_COL_PROFILESUMCASE].fillna(0)
            df = df.merge(df_ward_cluster_sum, how="left", left_on=ACC.CONST_COL_WARDID, right_on=ACC.CONST_COL_WARDID,suffixes=("","_SUMW"))
            df[ANXC_CONST_COL_WARDSUMCASE] = df[ANXC_CONST_COL_WARDSUMCASE].fillna(0)
            
            dict_col_display = {ACC.CONST_COL_WARDID:"Ward ID",ANXC_CONST_COL_WARDSUMCASE:"Total no. of\nobserved cases\nin any cluster by ward","profile_ID":"AMR Profile",
                                                   ACC.CONST_COL_NEWSDATE:"Start date",ACC.CONST_COL_NEWEDATE:"End date",
                                                   ACC.CONST_COL_NEWOBS:"No. of\nobserved cases\nin each cluster",ACC.CONST_COL_NEWPVAL:"P-value"}
            lst_col_display = list(dict_col_display.keys())
            
            df = prepare_tabletoplot(logger,df,lst_col=lst_col_display,lst_tosort=[ANXC_CONST_COL_WARDSUMCASE,ANXC_CONST_COL_DAYTOSTARTC,ACC.CONST_COL_NEWOBS], list_ascending=[False,True,False],dict_displaycol=dict_col_display,
                                     col_checkgroup=ACC.CONST_COL_WARDID,col_tohidevalue=ANXC_CONST_COL_WARDSUMCASE)
            
            lst_col_tosave = df.columns.tolist()
            df_resultdata_forsave = df.copy(deep=True)
            df_resultdata_forsave['Pathogen'] = sh_org
            df_resultdata_forsave['Specimen'] = spec
            df_resultdata_forsave = df_resultdata_forsave[['Pathogen','Specimen'] + lst_col_tosave]
            REP_AL.report_context(canvas_rpt, [shaveclustertext], 
               1.0*inch, 4.15*inch, 460, 50, font_size=11, font_align=TA_LEFT)
            #First up-to 20 clusters
            ioffset = 0.25*(len(df[:ANXC_CONST_CLUSTERLIST_MAXROW_FIRSTPAGE]) + 1)
            try:
                df1 = df[:ANXC_CONST_CLUSTERLIST_MAXROW_FIRSTPAGE]
                lst_df = [df1.columns.tolist()] + df1.values.tolist()
                table_draw = annexc_table_main(lst_df)
                table_draw.wrapOn(canvas_rpt, 480+ioffset, 300)
                table_draw.drawOn(canvas_rpt, 1.1*inch, (4.0-ioffset)*inch)
            except Exception as e:
                AL.printlog("Warning : generate first page for " + str(spec) + " of " + str(sh_org),False,logger) 
                logger.exception(e)
            #drop first 20 (ANXC_CONST_CLUSTERLIST_MAXROW_FISTPAGE)
            REP_AL.report_context(canvas_rpt,footnote_annexC(spec), 1.0*inch, 0.30*inch, 460, 100, font_size=9,line_space=12)
            REP_AL.canvas_printpage(canvas_rpt,page,lastpage,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,'Annex C',totalpage,startpage)
            df = df[ANXC_CONST_CLUSTERLIST_MAXROW_FIRSTPAGE:]
            page = page + 1
            if len(df)>0:
                #print(len(df))
                #cluster 21 - N
                imorepage = math.ceil(len(df)/ANXC_CONST_CLUSTERLIST_MAXROW_SECONDPAGEON)
                for i in range(imorepage):
                    #ipage = i + 1
                    REP_AL.report_title(canvas_rpt,ANNEXC_RPT_CONST_TITLE,1.07*inch, 10.5*inch,'#3e4444',font_size=16)
                    REP_AL.report_context(canvas_rpt, ["<b>"+ sspecname + ": " + sorgname + " (Continue)</b>"],                                       
                                   1.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
                    REP_AL.report_context(canvas_rpt, ["<b>"+ "Hospital-origin" + "</b>"],                                       
                                   6.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
                    REP_AL.report_context(canvas_rpt, [shaveclustertext], 
                                          1.0*inch, 9.0*inch, 460, 50, font_size=11, font_align=TA_LEFT)
                    ioffset = 0.25*(len(df[:ANXC_CONST_CLUSTERLIST_MAXROW_SECONDPAGEON]) + 1)
                    try:
                        df1 = df[:ANXC_CONST_CLUSTERLIST_MAXROW_SECONDPAGEON]
                        lst_df = [df1.columns.tolist()] + df1.values.tolist()
                        table_draw = annexc_table_main(lst_df)
                        table_draw.wrapOn(canvas_rpt, 480+ioffset, 300)
                        table_draw.drawOn(canvas_rpt, 1.0*inch, (8.9-ioffset)*inch)
                    except Exception as e:
                        AL.printlog("Warning : generate continue page for " + str(spec) + " of " + str(sh_org),False,logger) 
                        logger.exception(e)
                    #drop another 40 (ANXC_CONST_CLUSTERLIST_MAXROW_SECONDPAGEON)
                    REP_AL.report_context(canvas_rpt,footnote_annexC(spec), 1.1*inch, 0.30*inch, 460, 100, font_size=9,line_space=12)
                    REP_AL.canvas_printpage(canvas_rpt,page,lastpage,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,'Annex C',totalpage,startpage)
                    df = df[ANXC_CONST_CLUSTERLIST_MAXROW_SECONDPAGEON:]
                    page = page + 1
        except Exception as e:
            #print (e)
            AL.printlog("Warning : generate report for " + str(spec) + " of " + str(sh_org),False,logger) 
            logger.exception(e)
    else:
        REP_AL.report_context(canvas_rpt, [snoclustertext], 
               1.0*inch, 4.15*inch, 460, 50, font_size=11, font_align=TA_LEFT)
        REP_AL.report_context(canvas_rpt,footnote_annexC(spec), 1.0*inch, 0.30*inch, 460, 100, font_size=9,line_space=12)
        REP_AL.canvas_printpage(canvas_rpt,page,lastpage,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,'Annex C',totalpage,startpage)
        page = page + 1
    return [page,df_general_info_forsave,df_resultdata_forsave]
#Preparing date from 01 Dec 2022 to 2022/12/01
def get_dateforprm(ldate=""):
    fmt_date = ""
    month = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05","Jun":"06",
             "Jul":"07","Aug":"08","Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
    for keys,values in month.items():
        if keys in ldate:
            map_date = ldate.replace(keys,values).split(" ")
            fmt_date = str(map_date[2])+"/"+str(map_date[1])+"/"+str(map_date[0])
            break
    return fmt_date
#Retrieve date from Report1.csv
def retrieve_startEndDate(filename="",col_datafile=ACC.CONST_COL_DATAFILE,col_param=ACC.CONST_COL_PARAM,val_datafile=ACC.CONST_VALUE_DATAFILE,val_sdate=ACC.CONST_VALUE_SDATE,val_edate=ACC.CONST_VALUE_EDATE,col_date=ACC.CONST_COL_DATE):
    df = pd.read_csv(filename)
    start_date = get_dateforprm(df.loc[(df[col_datafile]==val_datafile)&(df[col_param]==val_sdate),col_date].tolist()[0])
    end_date   = get_dateforprm(df.loc[(df[col_datafile]==val_datafile)&(df[col_param]==val_edate),col_date].tolist()[0])
    return start_date, end_date
def main_generatepdf(canvas_rpt,logger,df_micro_ward,startpage,lastpage,totalpage,strgendate):
    sec1_res_i = AC.CONST_FILENAME_sec1_res_i
    bOK = False
    sub_printprocmem("Appendix C load and reformat result",logger)
    #Getting evaluation study
    fmt_studydate= "%Y/%m/%d"
    evaluation_study = retrieve_startEndDate(filename    =AC.CONST_PATH_RESULT+ sec1_res_i,
                                                 col_datafile="Type_of_data_file",
                                                 col_param   ="Parameters",
                                                 val_datafile="overall_data",
                                                 val_sdate   ="Minimum_date",
                                                 val_edate   ="Maximum_date",
                                                 col_date    ="Values")
    s_studydate  = evaluation_study[0]
    e_studydate  = evaluation_study[1]

    #create blank dataframe for graph generation
    df_week = create_df_weekday(s_studydate  =s_studydate, 
                                             e_studydate  =e_studydate, 
                                             fmt_studydate=fmt_studydate, 
                                             col_yearweek =ANXC_CONST_BASELINE_NEWVAR_WEEK,
                                             col_sweekday =ACC.CONST_COL_SWEEKDAY)
    #load dictionary ward
    df_dict_ward = pd.DataFrame()
    if AL.checkxlsorcsv(AC.CONST_PATH_ROOT,"dictionary_for_wards"):
        try:
            df_dict_ward = AL.readxlsorcsv_noheader_forceencode(AC.CONST_PATH_ROOT,"dictionary_for_wards", [AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL,"WARDTYPE","REQ","EXPLAINATION"],"utf-8",logger)
            df_dict_ward = df_dict_ward[df_dict_ward[AC.CONST_DICTCOL_DATAVAL].str.strip() != ""]
            df_dict_ward = df_dict_ward[df_dict_ward[AC.CONST_DICTCOL_AMASS].str.startswith("ward_")]
            df_dict_ward = df_dict_ward[[AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL]]
            df_dict_ward[AC.CONST_DICTCOL_DATAVAL] = df_dict_ward[AC.CONST_DICTCOL_DATAVAL].str.encode('utf-8', errors='ignore')
            df_dict_ward[AC.CONST_DICTCOL_DATAVAL] = df_dict_ward[AC.CONST_DICTCOL_DATAVAL].str.decode('utf-8', errors='ignore')
            #print(df_dict_ward)
            df_micro_ward = df_micro_ward.merge(df_dict_ward,how="left", left_on=AC.CONST_NEWVARNAME_WARDCODE,right_on=AC.CONST_DICTCOL_AMASS,suffixes=("", "DICT"))
        except Exception as e:
            AL.printlog("Warning : Unable to load ward dictionary: " +  str(e),False,logger)
            logger.exception(e) 
    else:
        AL.printlog("Note : No dictionary for ward file",False,logger)  
    
    #------------------------------------------------------------------------------------------------
    list_color = get_colorlist()
    #Load data and assign profile color
    dict_df_org = {} # key is org value is dict_df_sp which contains cluster df of that org for each specimen
    dict_df_org_baseline = {} # key is org value is dict_df_sp which contains baseline df (SatSCanInput)  of that org for each specimen
    dict_df_org_baseline_sum = {} # key is org value is dict_df_sp which contains baseline df summary of that org for each specimen
    dict_df_org_profile_sum = {}
    dict_df_org_profile_config = {}
    dict_org_profilecolor = {} # key is org value is dict of profile and it color assigned -> for cluster plus x profile with most cases
    dict_org_profilecolor_onlycluster = {} # key is org value is dict of profile and it color assigned
    dict_list_org_profile_sortbycase = {}
    for sh_org in ACC.dict_org.keys():
        df_pc= pd.DataFrame()
        try:
            df_pc= pd.read_excel(AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_PROFILE+"_"+ str.upper(sh_org) +".xlsx")
        except Exception as e:
            AL.printlog("Warning : Unable to profile information at: "+ AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_PROFILE+"_"+ str.upper(sh_org) +".xlsx" + " for " + str(sh_org),True,logger)
            logger.exception(e)
        dict_df_org_profile_config[sh_org] = df_pc
        #loop for load data to df and prepare dict color for all profile matching p-value condition
        #Global use for 
        dict_df_sp = {}
        dict_df_sp_baseline = {}
        dict_df_sp_baseline_sum = {}
        dict_df_sp_profile_sum = {}
        dict_profilecolor = {}
        dict_profilecolor_onlycluster = {}
        dict_list_sp_profile_sortbycase = {}
        #Local use
        list_cluster_profile = []
        list_top_num_case_profile = []
        list_profile_withcolor = []
        for sp in ACC.dict_spc.keys():
            #AL.printlog("Load data and reformat data for :  " + str(sp) + " of " + str(sh_org) ,False,logger) 
            #Variable persp
            df = pd.DataFrame()
            #c_df_profile_sum = pd.DataFrame()
            temp_df = pd.DataFrame()
            temp_df_sum = pd.DataFrame()
            temp_df_profile_sum = pd.DataFrame()
            list_profile_sortbycases = []
            #CLuster (AnnexC_listofpassedclusters_xxx_xxx.csv)
            sreplaceblankprofile = ANXC_CONST_BLANK_PROFILE_TO_PREFIX + str.upper(sh_org) + ANXC_CONST_BLANK_PROFILE_TO_SUFFIX
            sclusterfile = AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_PCLUSTER+"_"+ str.upper(sh_org)+"_"+ str(sp) +".xlsx"
            try:
                #df= pd.read_excel(AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_PCLUSTER+"_"+ str.upper(sh_org)+"_"+ str(sp) +".xlsx")
                df= pd.read_excel(sclusterfile)
                #Assign blank profile ID to profile such as profile_xxx_00
                df.loc[df[ACC.CONST_COL_PROFILEID] == "",ACC.CONST_COL_PROFILEID] = sreplaceblankprofile
                list_cluster_profile=list_cluster_profile + df[ACC.CONST_COL_PROFILEID].to_list()
                df[ACC.CONST_COL_WARDPROFILE] = df[ACC.CONST_COL_WARDID].str.strip() + ";" + df[ACC.CONST_COL_PROFILEID].str.strip()
                df[ANXC_CONST_COL_PROFILE_WITHCLUSTER] = df[ACC.CONST_COL_PROFILEID]
                df= AL.fn_clean_date(df,ACC.CONST_COL_NEWSDATE,ANXC_CONST_COL_STARTCDATE,"",logger)
                df= AL.fn_clean_date(df,ACC.CONST_COL_NEWEDATE,ANXC_CONST_COL_ENDCDATE,"",logger)
                df[ANXC_CONST_COL_DAYTOSTARTC] = 0
                df[ANXC_CONST_COL_DAYTOENDC] = 0
                try:
                    df[ANXC_CONST_COL_DAYTOSTARTC] =caldays(df, ANXC_CONST_COL_STARTCDATE, AC.CONST_ORIGIN_DATE)
                    df[ANXC_CONST_COL_DAYTOENDC] =caldays(df, ANXC_CONST_COL_ENDCDATE, AC.CONST_ORIGIN_DATE)
                except Exception as e:
                    AL.printlog("Warning : Unable to convert cluster signal start/end date: "+ str.upper(sh_org) + ":" + str(sp),True,logger)
                    #AL.printlog("Error : checkpoint Annex C generate report (load cluster data (Date convert) :  " + str(sp) + ") of " + str(sh_org) + " : " +str(e),True,logger) 
                    logger.exception(e)
                #c_df_profile_sum = df.groupby([ACC.CONST_COL_PROFILEID])[ACC.CONST_COL_CASE].sum().reset_index(name=ACC.CONST_COL_CASE)
                #c_df_profile_sum.rename(columns={ACC.CONST_COL_CASE:ANXC_CONST_COL_PROFILESUMCASE}, inplace=True)
                #c_df_profile_sum = c_df_profile_sum.sort_values(by=ANXC_CONST_COL_PROFILESUMCASE,ascending=False)
            except Exception as e:
                AL.printlog("Warning : Unable to load/reformat cluster signal at: "+ sclusterfile + " for " + str(sh_org) + ":" + str(sp) ,True,logger)
                logger.exception(e)
            dict_df_sp[sp] = df
            #Baseline (SatSCAN input file)
            #sbaselinefile = AC.CONST_PATH_TEMPWITH_PID, ACC.CONST_FILENAME_INPUT+"_" + str(sh_org) + "_" + str(sp),logger
            sbaselinefile =AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_INPUT+"_" + str(sh_org) + "_" + str(sp) +".xlsx"
            
            try:
                temp_df = AL.readxlsxorcsv(AC.CONST_PATH_TEMPWITH_PID, ACC.CONST_FILENAME_INPUT+"_" + str(sh_org) + "_" + str(sp),logger)
                #temp_df = pd.read_excel(sbaselinefile)
                temp_df = AL.fn_clean_date(temp_df,AC.CONST_NEWVARNAME_CLEANSPECDATE,ANXC_CONST_BASELINE_NEWVAR_SPECDATE,"",logger)
                temp_df[ANXC_CONST_COL_DAYTOBASELINEDATE] = 0
                try:
                    temp_df[ANXC_CONST_COL_DAYTOBASELINEDATE] =caldays(temp_df,ANXC_CONST_BASELINE_NEWVAR_SPECDATE, AC.CONST_ORIGIN_DATE)
                except Exception as e:
                    AL.printlog("Warning : Unable to convert specimen date: "+ str(sh_org) + ":" + str(sp),True,logger)
                    #AL.printlog("Error : checkpoint Annex C generate report (load baseline data (date convert):  " + str(sp) + " of " + str(sh_org) + " : " +str(e),True,logger) 
                    logger.exception(e)
                temp_df[ACC.CONST_COL_CASE] = temp_df[ACC.CONST_COL_CASE].apply(replace_non_integer)
                temp_df[ACC.CONST_COL_WARDID] = temp_df[ACC.CONST_COL_TESTGROUP].str.split(ANXC_CONST_BASELINE_VAR_WARDPROFILE_SPLITER).str[ANXC_CONST_BASELINE_VAR_WARDPROFILE_SPLITID_FOR_WARD].str.strip()
                temp_df[ACC.CONST_COL_PROFILEID] = temp_df[ACC.CONST_COL_TESTGROUP].str.split(ANXC_CONST_BASELINE_VAR_WARDPROFILE_SPLITER).str[ANXC_CONST_BASELINE_VAR_WARDPROFILE_SPLITID_FOR_PROFILE].str.strip()
                #Assign blank profile ID to profile such as profile_xxx_00
                temp_df.loc[temp_df[ACC.CONST_COL_PROFILEID] == "",ACC.CONST_COL_PROFILEID] = sreplaceblankprofile
                
                temp_df[ANXC_CONST_BASELINE_NEWVAR_WEEK] = temp_df[ANXC_CONST_BASELINE_NEWVAR_SPECDATE]
                temp_df[ANXC_CONST_BASELINE_NEWVAR_WEEK] = temp_df[ANXC_CONST_BASELINE_NEWVAR_WEEK].apply(getyearweek)   
                temp_df[ACC.CONST_COL_WARDPROFILE] = temp_df[ACC.CONST_COL_WARDID] + ";" + temp_df[ACC.CONST_COL_PROFILEID] 
                lst_col_before_merge = temp_df.columns.to_list()
                #Merge with cluter data to mark baseline -> is cluster
                if len(df) > 0:
                    temp_df2 = temp_df.merge(df, how="left", left_on=ACC.CONST_COL_WARDPROFILE, right_on=ACC.CONST_COL_WARDPROFILE,suffixes=("","_CLUSTER"))
                    temp_df2 = temp_df2[(temp_df2[ANXC_CONST_COL_DAYTOBASELINEDATE]>=temp_df2[ANXC_CONST_COL_DAYTOSTARTC]) & (temp_df2[ANXC_CONST_COL_DAYTOBASELINEDATE]<=temp_df2[ANXC_CONST_COL_DAYTOENDC])]
                    temp_df2 = temp_df2[lst_col_before_merge]
                    temp_df2[ANXC_CONST_COL_GOTCLUSTER] = 1
                    temp_df[ANXC_CONST_COL_GOTCLUSTER]  = 0
                    temp_df = pd.concat([temp_df, temp_df2], axis=0, ignore_index=True)
                    temp_df = temp_df.groupby(lst_col_before_merge)[ANXC_CONST_COL_GOTCLUSTER].sum().reset_index(name=ANXC_CONST_COL_GOTCLUSTER)
                    temp_df[ANXC_CONST_COL_PROFILE_WITHCLUSTER] = temp_df[ACC.CONST_COL_PROFILEID]
                    temp_df.loc[temp_df[ANXC_CONST_COL_GOTCLUSTER] ==0, ANXC_CONST_COL_PROFILE_WITHCLUSTER] = ANXC_CONST_PROFILENAME_FORNOCLUSTER
                else:
                    temp_df[ANXC_CONST_COL_GOTCLUSTER]  = 0
                    temp_df[ANXC_CONST_COL_PROFILE_WITHCLUSTER] = ANXC_CONST_PROFILENAME_FORNOCLUSTER
                #Base line summary
                temp_df_sum = temp_df.groupby([ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID])[ACC.CONST_COL_CASE].sum().reset_index(name=ACC.CONST_COL_CASE)
                temp_df_sum = temp_df_sum.sort_values(by=ACC.CONST_COL_CASE,ascending=False)
                temp_df_profile_sum = temp_df_sum.groupby([ACC.CONST_COL_PROFILEID])[ACC.CONST_COL_CASE].sum().reset_index(name=ACC.CONST_COL_CASE)
                temp_df_profile_sum.rename(columns={ACC.CONST_COL_CASE:ANXC_CONST_COL_PROFILESUMCASE}, inplace=True)
                temp_df_profile_sum = temp_df_profile_sum.sort_values(by=ANXC_CONST_COL_PROFILESUMCASE,ascending=False)
                list_profile_sortbycases = temp_df_profile_sum[ACC.CONST_COL_PROFILEID].to_list()
                list_top_num_case_profile = temp_df_profile_sum.head(3)[ACC.CONST_COL_PROFILEID].to_list()
                temp_df = temp_df.merge(temp_df_profile_sum, how="left", left_on=ACC.CONST_COL_PROFILEID, right_on=ACC.CONST_COL_PROFILEID,suffixes=("","_SUM"))
                temp_df[ANXC_CONST_COL_PROFILESUMCASE] = temp_df[ANXC_CONST_COL_PROFILESUMCASE].fillna(0)
            except Exception as e:
                AL.printlog("Warning : Unable to load/reformat baseline data at: "+ sbaselinefile + " for " + str(sh_org) + ":" + str(sp),True,logger)
                logger.exception(e)
            dict_df_sp_baseline[sp] = temp_df
            dict_df_sp_baseline_sum[sp] = temp_df_sum
            dict_df_sp_profile_sum[sp] = temp_df_profile_sum
            dict_list_sp_profile_sortbycase[sp] = list_profile_sortbycases
        #build dict for assign profile color
        list_profile_withcolor = list_top_num_case_profile + list_cluster_profile
        list_profile_withcolor =list(dict.fromkeys(list_profile_withcolor))   
        dict_profilecolor = get_dict_profile_color(list_profile_withcolor, list_color)
        dict_profilecolor_onlycluster = {key: value for key, value in dict_profilecolor .items() if key in list_cluster_profile}
        #add no cluster color
        dict_profilecolor[ANXC_CONST_PROFILENAME_FORNOCLUSTER] = ANXC_CONST_NOCLUSTER_COLOR
        dict_profilecolor_onlycluster[ANXC_CONST_PROFILENAME_FORNOCLUSTER] = ANXC_CONST_NOCLUSTER_COLOR
        #put to org variable
        dict_df_org[sh_org] = dict_df_sp
        dict_df_org_baseline[sh_org] = dict_df_sp_baseline
        dict_df_org_baseline_sum[sh_org] = dict_df_sp_baseline_sum
        dict_df_org_profile_sum[sh_org] = dict_df_sp_profile_sum
        dict_list_org_profile_sortbycase[sh_org] = dict_list_sp_profile_sortbycase 
        dict_org_profilecolor[sh_org] = dict_profilecolor
        dict_org_profilecolor_onlycluster[sh_org] = dict_profilecolor_onlycluster
    
    #------------------------------------------------------------------------------------------------
    #Print main annexC
    sub_printprocmem("Appendix C generate report",logger)
    df_generalinfo_all = pd.DataFrame()
    df_forsave_all = pd.DataFrame()
    df_graph_all = pd.DataFrame()
    prapare_mainAnnexC_mainpage(canvas_rpt,logger,startpage,startpage,lastpage,totalpage,strgendate)
    page = startpage + 1
    for sh_org in ACC.dict_org.keys():
        for sp in ACC.dict_spc.keys():
            sub_printprocmem("Generate annex C report for: " + str(sh_org) + ":" + str(sp),logger)
            try:
                df_cluster=dict_df_org[sh_org][sp]
                df_baseline = dict_df_org_baseline[sh_org][sp]
                df_baseline_sum = dict_df_org_baseline_sum[sh_org][sp]
                #df_profile_sum = dict_df_org_profile_sum[sh_org][sp]
                if len(df_baseline) > 0:
                    df_graph = df_week.copy(deep=True)
                    df_graph = df_graph.merge(df_baseline, how="left", left_on=ANXC_CONST_BASELINE_NEWVAR_WEEK, right_on=ANXC_CONST_BASELINE_NEWVAR_WEEK,suffixes=("", "_DUP"))
                    # copy data for further save to result data folder
                    df_graph_tosave = df_graph.copy(deep=True)
                    lst_col_tosave = [ACC.CONST_COL_SWEEKDAY,ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_CASE,ANXC_CONST_BASELINE_NEWVAR_SPECDATE,
                                      ANXC_CONST_BASELINE_NEWVAR_WEEK,ANXC_CONST_COL_PROFILE_WITHCLUSTER]
                    df_graph_tosave['Pathogen'] = sh_org
                    df_graph_tosave['Specimen'] = sp
                    df_graph_tosave = df_graph_tosave[['Pathogen','Specimen'] + lst_col_tosave]
                    df_graph_tosave[ANXC_CONST_COL_PROFILE_WITHCLUSTER] = df_graph_tosave[ANXC_CONST_COL_PROFILE_WITHCLUSTER].fillna(ANXC_CONST_PROFILENAME_FORNOCLUSTER)
                    df_graph_tosave[ACC.CONST_COL_CASE] = df_graph_tosave[ACC.CONST_COL_CASE].fillna(0)
                    df_graph_tosave.sort_values(by=[ANXC_CONST_BASELINE_NEWVAR_WEEK], ascending=[True])
                    #-------
                    df_graph = df_graph[[ACC.CONST_COL_SWEEKDAY,ANXC_CONST_BASELINE_NEWVAR_WEEK,ANXC_CONST_COL_PROFILE_WITHCLUSTER,ACC.CONST_COL_CASE]]
                    df_graph[ANXC_CONST_COL_PROFILE_WITHCLUSTER] = df_graph[ANXC_CONST_COL_PROFILE_WITHCLUSTER].fillna(ANXC_CONST_PROFILENAME_FORNOCLUSTER)
                    df_graph[ACC.CONST_COL_CASE] = df_graph[ACC.CONST_COL_CASE].fillna(0)
                    gen_cluster_graph(logger,df=df_graph,
                                      xcol_sort=ANXC_CONST_BASELINE_NEWVAR_WEEK,
                                      xcol_display=ACC.CONST_COL_SWEEKDAY,
                                      ycol=ACC.CONST_COL_CASE,
                                      zcol=ANXC_CONST_COL_PROFILE_WITHCLUSTER,
                                      list_privotcol=dict_list_org_profile_sortbycase[sh_org][ANXC_CONST_SP_ALL],
                                      #zcol_descsort=scol_profilesumvase_onlycluster,
                                      dict_profile_color=dict_org_profilecolor_onlycluster[sh_org], 
                                      filename=AC.CONST_PATH_TEMPWITH_PID + ANXC_CONST_CLUSTER_GRAPH_FNAME+"_" + str(sh_org) + "_" + str(sp) + ".png",
                                      figsizex=20,figsizey=10)
                    try:
                        if len(df_graph_tosave) > 0:
                            if len(df_graph_all) >0:
                                df_graph_all = pd.concat([df_graph_all, df_graph_tosave], ignore_index=True)
                            else:
                                df_graph_all = df_graph_tosave.copy(deep=True)
                    except Exception as e:
                        AL.printlog("Warning : Unable to generate profile graph data for export to resultdata folder: "+ str(sh_org) + ":" + str(sp),True,logger)
                        logger.exception(e)
                resultlist = prapare_mainAnnexC_per_org(canvas_rpt,logger,page,startpage,lastpage,totalpage,strgendate,pvaluelimit=ANNEXC_RPT_CONST_PVALUELIMIT,
                                           sh_org=sh_org,spec=sp,
                                           df=df_cluster, 
                                           df_baseline = df_baseline,
                                           df_profile_sum = dict_df_org_profile_sum[sh_org][ANXC_CONST_SP_ALL],
                                           filename_graph=AC.CONST_PATH_TEMPWITH_PID + ANXC_CONST_CLUSTER_GRAPH_FNAME+"_" + str(sh_org) + "_" + str(sp) + ".png")
                page = resultlist[0]
                try:
                    df_forsave = resultlist[1]
                    if len(df_forsave) > 0:
                        if len(df_generalinfo_all) >0:
                            df_generalinfo_all = pd.concat([df_generalinfo_all, df_forsave], ignore_index=True)
                        else:
                            df_generalinfo_all = df_forsave.copy(deep=True)
                except Exception as e:
                    AL.printlog("Warning : Unable to generate genral information data for export to resultdata folder: "+ str(sh_org) + ":" + str(sp),True,logger)
                    logger.exception(e)
                try:
                    df_forsave = resultlist[2]
                    if len(df_forsave) > 0:
                        if len(df_forsave_all) >0:
                            df_forsave_all = pd.concat([df_forsave_all, df_forsave], ignore_index=True)
                        else:
                            df_forsave_all = df_forsave.copy(deep=True)
                except Exception as e:
                    AL.printlog("Warning : Unable to generate cluster table data for export to resultdata folder: "+ str(sh_org) + ":" + str(sp),True,logger)
                    logger.exception(e)
            except Exception as e:
                AL.printlog("Warning : Unable to generate report: "+ str(sh_org) + ":" + str(sp),True,logger)
                logger.exception(e)
    if not AL.fn_savexlsx(df_generalinfo_all, AC.CONST_PATH_RESULT + "ver3_AnnexC_generalinfo.xlsx", logger):
        AL.printlog("Warning : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "ver3_AnnexC_generalinfo.xlsx",False,logger) 
    if not AL.fn_savexlsx(df_forsave_all, AC.CONST_PATH_RESULT + "ver3_AnnexC_wardprofile_cluster.xlsx", logger):
        AL.printlog("Warning : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "ver3_AnnexC_wardprofile_cluster.xlsx",False,logger)  
    if not AL.fn_savexlsx(df_graph_all, AC.CONST_PATH_RESULT + "ver3_AnnexC_profilegraph.xlsx", logger):
        AL.printlog("Warning : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "ver3_AnnexC_profilegraph.xlsx",False,logger)  
    #------------------------------------------------------------------------------------------------
    #Print main supplementary annexC
    sub_printprocmem("Appendix C generate supplementary report",logger)
    path_input = AC.CONST_PATH_ROOT
    path_iden = AC.CONST_PATH_REPORTWITH_PID
    
    canvas_sup_rpt = canvas.Canvas(path_iden +"Supplementary_data_Annex_C.pdf")
    try:
        cover(canvas_sup_rpt,logger,strgendate)
    except Exception as e:
        AL.printlog("Warning : Unable to generate cover page: "+ str(sh_org) + ":" + str(sp),True,logger)
        logger.exception(e)
        pass
    totalpage = 999 #Need confirm if not need to display total page in supplementary report that this will be remove (other part in the code too)
    lastpage = totalpage #Need confirm if not need to display total page in supplementary report that this will be remove (other part in the code too)
    startpage = 1
    page = 1
    for sh_org in ACC.dict_org.keys():
        df_pc = dict_df_org_profile_config[sh_org]
        df_pc = df_pc.iloc[:, :-2]
        #print(df_pc)
        list_profile_atb_column = df_pc.columns.to_list()
        list_profile_atb_column = [x for x in list_profile_atb_column if x != ACC.CONST_COL_PROFILEID]
        for sp in ACC.dict_spc.keys():
            sub_printprocmem("Generate annex C supplementary report for: " + sh_org + ":" + sp,logger)
            try:
                df_cluster=dict_df_org[sh_org][sp]
                df_baseline = dict_df_org_baseline[sh_org][sp]
                df_baseline_sum = dict_df_org_baseline_sum[sh_org][sp]
                
                df_profile_sum = dict_df_org_profile_sum[sh_org][sp]
                df_profile_sum_det= pd.DataFrame()
                if len(df_profile_sum) > 0:
                    try:
                        df_profile_sum_det = df_profile_sum.merge(df_pc, how="left", left_on=ACC.CONST_COL_PROFILEID, right_on=ACC.CONST_COL_PROFILEID,suffixes=("", "_DUP"))
                        df_profile_sum_det = df_profile_sum_det[[ACC.CONST_COL_PROFILEID] +list_profile_atb_column + [ANXC_CONST_COL_PROFILESUMCASE]]
                    except Exception as e:
                         AL.printlog("Warning : Unable to calculate summary by profile: "+ str(sh_org) + ":" + str(sp),True,logger)
                         logger.exception(e)   
                df_ward_sum = pd.DataFrame()
                if len(df_baseline) > 0:
                    try:
                        df_ward_sum = df_baseline_sum.groupby([ACC.CONST_COL_WARDID])[ACC.CONST_COL_CASE].sum().reset_index(name=ACC.CONST_COL_CASE)
                        #print(df_ward_sum)
                        df_ward_sum = df_baseline.groupby([ACC.CONST_COL_WARDID])[ACC.CONST_COL_CASE,ANXC_CONST_COL_GOTCLUSTER].sum().reset_index()
                        df_ward_sum.rename(columns={ACC.CONST_COL_CASE:ANXC_CONST_COL_WARDSUMCASE}, inplace=True)
                        df_ward_sum = df_ward_sum.sort_values(by=ANXC_CONST_COL_WARDSUMCASE,ascending=False).reset_index()
                        df_ward_sum[ANXC_CONST_COL_SUP_DISPLAYWARDGRAPH] = 0
                        if ANXC_CONST_NUM_TOPWARD_TODISPLAY > 0:
                            if len(df_ward_sum) < ANXC_CONST_NUM_TOPWARD_TODISPLAY:
                                df_ward_sum[ANXC_CONST_COL_SUP_DISPLAYWARDGRAPH] = 1
                            else:
                                df_ward_sum.loc[df.head(ANXC_CONST_NUM_TOPWARD_TODISPLAY).index, ANXC_CONST_COL_SUP_DISPLAYWARDGRAPH] = 1
                            #df.loc[df[(df.A % 2 == 0)].head(10).index,'B'] = 100
                        df_ward_sum.loc[df_ward_sum[ANXC_CONST_COL_GOTCLUSTER] >0, ANXC_CONST_COL_SUP_DISPLAYWARDGRAPH] = 1
                        #Merge to get ward name
                        try:
                            lst_col_before_merge = df_ward_sum.columns.tolist()
                            df_ward_sum = df_ward_sum.merge(df_dict_ward,how="left", left_on=ACC.CONST_COL_WARDID, right_on=AC.CONST_DICTCOL_AMASS,suffixes=("", "_WD"))
                            df_ward_sum = df_ward_sum[lst_col_before_merge +[AC.CONST_DICTCOL_DATAVAL]]
                        except Exception as e:
                            AL.printlog("Warning : Unable to merge with ward dict: "+ str(sh_org) + ":" + str(sp),True,logger)
                            logger.exception(e)
                    except Exception as e:
                        AL.printlog("Warning : Unable to calculate summary by ward: "+ str(sh_org) + ":" + str(sp),True,logger)
                        logger.exception(e)
                if len(df_ward_sum) > 0:
                    #df_ward_sum = df_ward_sum.head(2)
                    #itopdis = 0
                    for i in range(len(df_ward_sum)):
                        try:
                            swardid = df_ward_sum.iloc[i][ACC.CONST_COL_WARDID]
                            temp_df = df_baseline[df_baseline[ACC.CONST_COL_WARDID] == swardid]
                            if df_ward_sum.iloc[i][ANXC_CONST_COL_SUP_DISPLAYWARDGRAPH] > 0:
                                df_graph = df_week.copy(deep=True)
                                df_graph = df_graph.merge(temp_df, how="left", left_on=ANXC_CONST_BASELINE_NEWVAR_WEEK, right_on=ANXC_CONST_BASELINE_NEWVAR_WEEK,suffixes=("", "_DUP"))
                                df_graph = df_graph[[ACC.CONST_COL_SWEEKDAY,ANXC_CONST_BASELINE_NEWVAR_WEEK,ANXC_CONST_COL_PROFILE_WITHCLUSTER,ACC.CONST_COL_CASE]]
                                df_graph[ANXC_CONST_COL_PROFILE_WITHCLUSTER] = df_graph[ANXC_CONST_COL_PROFILE_WITHCLUSTER].fillna(ANXC_CONST_PROFILENAME_FORNOCLUSTER)
                                df_graph[ACC.CONST_COL_CASE] = df_graph[ACC.CONST_COL_CASE].fillna(0)
                                gen_cluster_graph(logger,df=df_graph,
                                                  xcol_sort=ANXC_CONST_BASELINE_NEWVAR_WEEK,
                                                  xcol_display=ACC.CONST_COL_SWEEKDAY,
                                                  ycol=ACC.CONST_COL_CASE,
                                                  zcol=ANXC_CONST_COL_PROFILE_WITHCLUSTER,
                                                  list_privotcol=dict_list_org_profile_sortbycase[sh_org][ANXC_CONST_SP_ALL],
                                                  #zcol_descsort=scol_profilesumvase_onlycluster,
                                                  dict_profile_color=dict_org_profilecolor_onlycluster[sh_org], 
                                                  filename=AC.CONST_PATH_TEMPWITH_PID + ANXC_CONST_BASELINE_GRAPH_FNAME + "_" + str(sh_org) + "_" + str(sp) + "_" + str(swardid) + ".png",
                                                  figsizex=25,figsizey=5)
                            #itopdis = itopdis + 1
                        except Exception as e:
                            AL.printlog("Warning : Unable to generate graph: "+ str(sh_org) + ":" + str(sp),True,logger)
                            logger.exception(e)   
                #AL.printlog("Finish generating base line graph per organism : " + str(sp) + " of " + str(sh_org) ,False,logger)
                
                page = prapare_supplementAnnexC_per_org(canvas_sup_rpt,logger,page,startpage,lastpage,totalpage,strgendate,
                                                        sh_org=sh_org,spec=sp,
                                                        df=df_cluster,
                                                        df_baseline=df_baseline,
                                                        df_profile=df_profile_sum_det,
                                                        df_ward=df_ward_sum,
                                                        list_profile_atb_column=list_profile_atb_column)
            except Exception as e:
                AL.printlog("Warning : Unable to generate report: " + str(sh_org) + ":" + str(sp),True,logger)
                logger.exception(e)
                pass
   
    #------------------------------------------------------------
    #list of ward
    lst_microward_footnote =["* In case that there are ward names in your hospital_admission_data file, this list and the analysis will prioritize the ward names in the microbiology_data file over the ones in hospital_admission_data file. "]
    #print(df_micro_ward)
    if len(df_micro_ward)>0:
        df = df_micro_ward[[AC.CONST_NEWVARNAME_WARDCODE,AC.CONST_DICTCOL_DATAVAL]]
        df.rename(columns={AC.CONST_NEWVARNAME_WARDCODE:"Ward ID",AC.CONST_DICTCOL_DATAVAL:"Ward name"}, inplace=True)
        imaxcol = 1
        icurcol = 1
        icoloffset = 5
        balreadyprintpage = False
        bisfirstpage = True
        while len(df) > 0:
            if icurcol == 1:
                #print header
                balreadyprintpage = False
                REP_AL.report_title(canvas_sup_rpt,"Table S1: List of ward names in your microbiology_data file",1.07*inch, 10.6*inch,'#3e4444',font_size=16)
                scon = " (Continue)" if bisfirstpage == False else ""
                REP_AL.report_context(canvas_sup_rpt, ["<b>"+ "List of ward" + scon + "</b>"],                                       
                               1.0*inch, 9.5*inch, 460, 50, font_size=13, font_align=TA_LEFT)
            df1 = df[:ANXC_CONST_MAX_BASELINETBLROW_PERPAGE]
            ioffset = 0.25*(len(df1) + 1)
            try:
                lst_df = [df1.columns.tolist()] + df1.values.tolist()
                table_draw = annexc_table_sup_micro_ward(lst_df,[1*inch,4*inch])
                table_draw.wrapOn(canvas_sup_rpt, 480+ioffset, 300)
                table_draw.drawOn(canvas_sup_rpt, (1+((icurcol-1)*icoloffset))*inch, (9.3-ioffset)*inch)
            except Exception as e:
                AL.printlog("Error : gen micro ward table",True,logger) 
                logger.exception(e)
            df = df[ANXC_CONST_MAX_BASELINETBLROW_PERPAGE:]
            if icurcol >= imaxcol-1:
                if balreadyprintpage == False:
                    REP_AL.report_context(canvas_sup_rpt,lst_microward_footnote, 1.0*inch, 0.30*inch, 460, 70, font_size=9,line_space=12)
                    REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
                    bisfirstpage = False
                    page = page + 1
                    icurcol = 1
                    balreadyprintpage = True    
            else:
                icurcol = icurcol + 1
        if balreadyprintpage == False:
            REP_AL.report_context(canvas_sup_rpt,lst_microward_footnote, 1.0*inch, 0.30*inch, 460, 70, font_size=9,line_space=12)
            REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
            bisfirstpage = False
            page = page + 1
            icurcol = 1
            balreadyprintpage = True
    else:
        REP_AL.report_title(canvas_sup_rpt,"Table S1: List of ward names in your microbiology_data file",1.07*inch, 10.6*inch,'#3e4444',font_size=16)
        REP_AL.report_context(canvas_sup_rpt, ["None"],
                       1.5*inch, 9.2*inch, 460, 50, font_size=11, font_align=TA_LEFT)
        REP_AL.report_context(canvas_sup_rpt,lst_microward_footnote, 1.0*inch, 0.30*inch, 460, 70, font_size=9,line_space=12)
        REP_AL.canvas_printpage_nototalpage(canvas_sup_rpt,page,strgendate,True,AC.CONST_REPORTPAGENUM_MODE,ANXC_CONST_FOOT_REPNAME)
        page = page + 1
    canvas_sup_rpt.save()
    bOK = True
    return bOK
#Call by AMASS core to run generate annex C
def generate_annex_c_report(canvas_rpt,logger,df_micro_ward,startpage,lastpage,totalpage,strgendate):
    #AL.printlog("AMR surveillance report - checkpoint Annex C main page",False,logger)
    sub_printprocmem("AMR surveillance report - checkpoint annex C",logger)
    if main_generatepdf(canvas_rpt,logger,df_micro_ward,startpage,lastpage,totalpage,strgendate) == True:    
        #Write success log
        pass
    else:
        AL.printlog("Error : checkpoint Annex C for each organism",False,logger)    