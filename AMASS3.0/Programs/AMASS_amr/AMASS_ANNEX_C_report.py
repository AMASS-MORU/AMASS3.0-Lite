import pandas as pd
from reportlab.pdfgen import canvas #creating PDF page
from reportlab.lib.units import inch #importing inch for plotting
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY #setting paragraph style
#import AMASS_amr_const as AC
import AMASS_ANNEX_C_const as ACC
import AMASS_ANNEX_C_commonlib as ALC
import AMASS_amr_commonlib as AL
import AMASS_amr_commonlib_report as REP_AL

#----------------------------------------------------------------------------------------------------------
#Prepare data functions 
#Assigning number of patient and number of wards
def assign_valuetodf(df=pd.DataFrame(),idx_assign="",col_assign="",val_toassign=0):
    try:
        df.at[idx_assign,col_assign] = val_toassign
    except Exception as e:
        print (e)
    return df
#Reordering clusters by 1.total no.observed cases by wards and 2.start signal date
#return a dataframe with ordered clusters
def reorder_clusters(df=pd.DataFrame(),
                     col_wardid=ACC.CONST_COL_WARDID,col_obs=ACC.CONST_COL_OBS,col_date=ACC.CONST_COL_CLEANSDATE):
    df_final = pd.DataFrame()
    df_1     = df.loc[:,[col_wardid,col_obs]]
    df_1.loc[:,col_obs] = df_1.loc[:,col_obs].astype(int)
    lst_ward = df_1.groupby(by=[col_wardid]).sum().sort_values(by=[col_obs],ascending=False).index.tolist()
    for ward in lst_ward:
        df_temp  = df.loc[df[col_wardid]==ward,:].sort_values(by=[col_date,col_obs],ascending=[True,True],na_position="last")
        df_final = pd.concat([df_final,df_temp],axis=0)
    return df_final



#Main function for prepare data
def main_preparedata(logger):
    bOK = False
    """
    evaluation_study = ALC.retrieve_startEndDate(filename    =AC.CONST_PATH_RESULT+ACC.CONST_FILENAME_REPORT1,
                                                 col_datafile="Type_of_data_file",
                                                 col_param   ="Parameters",
                                                 val_datafile="microbiology_data",
                                                 val_sdate   ="Minimum_date",
                                                 val_edate   ="Maximum_date",
                                                 col_date    ="Values")
    """
    
    #s_studydate  = evaluation_study[0]
    #e_studydate  = evaluation_study[1]
    #b_wardhighpat = True
    #num_wardhighpat = 1
    
    for sh_spc in ACC.dict_spc.keys():
        #df_num = pd.DataFrame(0,index=ACC.dict_org.keys(),columns=[ACC.CONST_COL_CASE,ACC.CONST_COL_WARDID])
        for sh_org in ACC.dict_org.keys():
            #lo_org = ACC.dict_org[sh_org][0]
            #fmt_studydate= "%Y/%m/%d"
            filename_result = ACC.CONST_FILENAME_RESULT+"_"+sh_org+"_"+sh_spc+".col.txt"
            #filename_input  = ACC.CONST_FILENAME_INPUT +"_"+sh_org+"_"+sh_spc+".csv"
    
            #create blank dataframe for graph generation
            """
            df_graph = ALC.create_df_forstartweekday(s_studydate  =s_studydate, 
                                                     e_studydate  =e_studydate, 
                                                     fmt_studydate=fmt_studydate, 
                                                     col_sweekday =ACC.CONST_COL_SWEEKDAY)
            """
            #read satscan result
            df_re = pd.DataFrame()
            try:
                df_re = ALC.read_texttocol(ACC.CONST_PATH_ANNEXC_RESULT+filename_result)
                #print(ACC.CONST_PATH_ANNEXC_RESULT+filename_result)
                #ward_015;Profile_7;crab;ward_model
                df_re[[ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_RESISTPROFILE]] = df_re.loc[:,ACC.CONST_COL_LOCID].str.split(";", expand=True)
            except Exception as e:
                #print (e)
                AL.printlog("Error : checkpoint Annex C prepare data of " + str(sh_org) + " : " +str(e),True,logger) 
                logger.exception(e)
                df_re[[ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_RESISTPROFILE]] = ""
    
            #formatting datetime
            if len(df_re) > 0:
                df_re_date = AL.fn_clean_date(df=df_re,     oldfield=ACC.CONST_COL_SDATE,cleanfield=ACC.CONST_COL_CLEANSDATE,dformat="",logger="")
                df_re_date = AL.fn_clean_date(df=df_re_date,oldfield=ACC.CONST_COL_EDATE,cleanfield=ACC.CONST_COL_CLEANEDATE,dformat="",logger="")
                df_re_date[ACC.CONST_COL_CLEANSDATE] = df_re_date[ACC.CONST_COL_CLEANSDATE].astype(str)
                df_re_date[ACC.CONST_COL_CLEANEDATE] = df_re_date[ACC.CONST_COL_CLEANEDATE].astype(str)
                #df_re_date_pval = ALC.format_lancent_pvalue(df=df_re_date,col_pval=ACC.CONST_COL_PVAL,col_clean_pval=ACC.CONST_COL_CLEANPVAL)
            else:
                df_re_date = df_re.copy()
                df_re_date[[ACC.CONST_COL_CLEANPVAL,ACC.CONST_COL_CLEANSDATE,ACC.CONST_COL_CLEANEDATE]] = ""
    
            #preparing date and export results
            select_col = [ACC.CONST_COL_ORGNAME,ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_CLEANSDATE,ACC.CONST_COL_CLEANEDATE,
                          ACC.CONST_COL_OBS,ACC.CONST_COL_CLEANPVAL,ACC.CONST_COL_CLEANPVAL+"_lancent"]
            try:
                df_re_date_1 = df_re_date.copy()
                df_re_date_1[ACC.CONST_COL_ORGNAME]   = df_re_date_1[ACC.CONST_COL_RESISTPROFILE].map(ACC.dict_org)
                df_re_date_1[[ACC.CONST_COL_ORGNAME]] = df_re_date_1[ACC.CONST_COL_ORGNAME][0][1]
                df_re_date_1 = df_re_date_1.loc[:,select_col]
                df_re_date_2 = df_re_date_1.loc[df_re_date_1[ACC.CONST_COL_CLEANPVAL].astype(float)<=0.05,:]
                df_re_date_2[ACC.CONST_COL_WARDID] = df_re_date_2[ACC.CONST_COL_WARDID]+"*"
                reorder_clusters(df=pd.concat([df_re_date_2,df_re_date_1.loc[df_re_date_1[ACC.CONST_COL_CLEANPVAL].astype(float)>0.05,:]])).rename(columns={ACC.CONST_COL_WARDID:ACC.CONST_COL_WARDID,
                                                                                                                          ACC.CONST_COL_CLEANSDATE:ACC.CONST_COL_NEWSDATE,
                                                                                                                          ACC.CONST_COL_CLEANEDATE:ACC.CONST_COL_NEWEDATE,
                                                                                                                          ACC.CONST_COL_OBS:ACC.CONST_COL_NEWOBS, 
                                                                                                                          ACC.CONST_COL_CLEANPVAL:ACC.CONST_COL_NEWPVAL,
                                                                                                                          ACC.CONST_COL_CLEANPVAL+"_lancent":ACC.CONST_COL_NEWPVAL+"lancent"}).to_excel(ACC.CONST_PATH_ANNEXC_RESULT+ACC.CONST_FILENAME_ACLUSTER+"_"+str(sh_org)+"_"+str(sh_spc)+".xlsx",index=False,header=True)
                ALC.retrieve_databycondition(df=reorder_clusters(df=df_re_date_1),col_pval=ACC.CONST_COL_CLEANPVAL).rename(columns={ACC.CONST_COL_WARDID:ACC.CONST_COL_WARDID,
                                                                                                               ACC.CONST_COL_CLEANSDATE:ACC.CONST_COL_NEWSDATE,
                                                                                                               ACC.CONST_COL_CLEANEDATE:ACC.CONST_COL_NEWEDATE,
                                                                                                               ACC.CONST_COL_OBS:ACC.CONST_COL_NEWOBS, 
                                                                                                               ACC.CONST_COL_CLEANPVAL+"_lancent":ACC.CONST_COL_NEWPVAL}).drop(columns=[ACC.CONST_COL_CLEANPVAL]).to_excel(ACC.CONST_PATH_ANNEXC_RESULT+ACC.CONST_FILENAME_PCLUSTER+"_"+str(sh_org)+"_"+str(sh_spc)+".xlsx",index=False,header=True)
            except Exception as e:
                #print (e)
                AL.printlog("Error : checkpoint Annex C prepare data of " + str(sh_org) + " : " +str(e),True,logger) 
                logger.exception(e)
                df_re_date_1 = pd.DataFrame(columns=select_col)
    bOK = True
    return bOK
#----------------------------------------------------------------------------------------------------------
# Generate ANNEX C report in PDF
def template_page1_intro():
    return ["This supplementary report shows the information of potential cluster which is frequently found within short period comparing to a baseline in a normal situation of each hospital. This report is generated by default, even if hospital_admission_data is unavailable. This is allow to utilize microbiology_data to de-duplicate and perform the prediction automatically.", 
            "<br/>" + "Please note that there are two models including ward model and hospital-wide model. The prediction is a tier-based analysis performing based on an provided microbiology_data and hospital_admission_data files. A ward model performs when user provides wards information in microbiology_data. If there is no wards information, the hospital-wide model is performed and all patients are assumed to occur in one site.",
            "And if hospital_admission_data is available, in-hospital patients for each pathogen are included for the prediction. "]
def template_page1_org():
    return ["Methicillin-NS <i>S. aureus</i> (MRSA)",
            "Vancomycin-NS <i>Enterococcus</i> spp. (VRE) for <i>E. faecalis</i> and <i>E. faecium</i>",
            "Carbapenem-NS Enterobacteriaceae (CRE) for <i>E. coli</i> and <i>K. pneumoniae</i>",
            "Carbapenem-NS <i>P. aeruginosa</i> (CRPA)",
            "Carbapenem-NS <i>A. baumannii</i> (CRAB)"]
def template_page1_result():
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    return ["The data included in the analysis had:"]
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
def footnote():
    return ["*AnnexC de-duplicated the data by including the first resistant isolate per patient per evaluation period. "+
            "Bar graphs show wards with two top cluster with highest number of patient and p-value<0.05 (or wards with top two highest number of patients, when no. of cluster is less than two). "+
            "NA=Not available"]

def prepare_tabletoplot(filename="",lst_tosort=[], ascending=False):
    lst_df = []
    try:
        df = pd.read_excel(filename)
        if len(df) >= 2:     #has at least 1 cluster with p-value <= 0.05 >>> sort number of cases
            df = df.sort_values(by=lst_tosort,ascending=ascending).iloc[:2,1:]
            lst_df = [df.columns.tolist()] + df.values.tolist()
        elif len(df) == 1:
            df = df.sort_values(by=lst_tosort,ascending=ascending).iloc[:2,1:]
            lst_df = [df.columns.tolist()] + df.values.tolist() + [["NA","NA","NA","NA","NA","NA"]]
        else:               #has at least 1 cluster with p-value > 0.05 >>> ignore
            pass
    except:                 #has no cluster >>> ignore
        pass
    return lst_df

def prapare_mainAnnexC_mainpage(canvas_rpt,logger,filename_pdf="",page_number=2):
    #c = canvas.Canvas(filename_pdf+"_"+str(1)+".pdf")
    REP_AL.report_title(canvas_rpt,"Annex C: Cluster signal identified by SaTScan method",1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    REP_AL.report_title(canvas_rpt,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    REP_AL.report_context(canvas_rpt, template_page1_intro(),                                           
                   1.0*inch, 6.6*inch, 460, 200, font_size=11, font_align=TA_LEFT)
    REP_AL.report_title(canvas_rpt,'Pathogens under the survey',1.07*inch, 6.2*inch,'#3e4444',font_size=12)
    REP_AL.report_context(canvas_rpt, template_page1_org(),                                           
                   1.0*inch, 4.0*inch, 460, 150, font_size=11, font_align=TA_LEFT)
    REP_AL.report_title(canvas_rpt,'Results',1.07*inch, 4.1*inch,'#3e4444',font_size=12)
    REP_AL.report_context(canvas_rpt, template_page1_result(),                                           
                   1.0*inch, 1.2*inch, 460, 200, font_size=11, font_align=TA_LEFT)
    REP_AL.report_todaypage(canvas_rpt,270,30,"AnnexC - Page " + str(1))
    canvas_rpt.showPage()
    #c.save()

def prapare_mainAnnexC_per_org(canvas_rpt,logger,df_all=pd.DataFrame(), df_blo=pd.DataFrame(), sh_org="",
                             num_pat_all=0, num_pat_blo=0, num_ward_all=0, num_ward_blo=0,
                             filename_pdf="", filename_graph="", page_number=2):
    
    #c = canvas.Canvas(filename_pdf+"_"+str(page_number)+".pdf")
    REP_AL.report_title(canvas_rpt,"Annex C: Cluster signal identified by SaTScan method",1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    REP_AL.report_context(canvas_rpt, ["<b>"+"Blood : "+ACC.dict_org[sh_org][2]+"</b>"],                                           
                   1.0*inch, 9.5*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    REP_AL.report_context(canvas_rpt, ["<b><i><font color=\"#000080\">"+"No of patients with HO "+ACC.dict_org[sh_org][2]+" BSI* : " + str(num_pat_blo)+"</font></i></b>"],              
                   1.0*inch, 9.3*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    REP_AL.report_context(canvas_rpt, ["<b><i><font color=\"#000080\">"+"No of wards having patients with "+ACC.dict_org[sh_org][2]+" BSI : " + str(num_ward_blo)+"</font></i></b>"], 
                   1.0*inch, 9.1*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    try:
        canvas_rpt.drawImage(filename_graph+"_"+sh_org+"_blo.png", 0.1*inch, 7.1*inch, preserveAspectRatio=True, width=8.0*inch, height=2.40*inch,showBoundary=False) 
    except Exception as e:
        AL.printlog("Error : checkpoint Annex C generate report (graph blood) of " + str(sh_org) + " : " +str(e),True,logger) 
        logger.exception(e)
        #print (e)
    if len(df_blo)>0:
        try:
            REP_AL.report_context(canvas_rpt, ["Table of the two top clusters with p-value<0.05"], 
               1.0*inch, 6.5*inch, 460, 50, font_size=11, font_align=TA_LEFT)
            table_draw = REP_AL.annexc_table_main(df_blo)
            table_draw.wrapOn(canvas_rpt, 500, 300)
            table_draw.drawOn(canvas_rpt, 1.5*inch, 5.9*inch)
        except Exception as e:
            #print (e)
            AL.printlog("Error : checkpoint Annex C generate report (table blood) of " + str(sh_org) + " : " +str(e),True,logger) 
            logger.exception(e)
    else:
        REP_AL.report_context(canvas_rpt, ["There is no cluster with p-value<0.05."], 
               1.0*inch, 6.3*inch, 460, 50, font_size=11, font_align=TA_LEFT)

    REP_AL.report_context(canvas_rpt, ["<b>"+"All specimens : "+ACC.dict_org[sh_org][2]+"</b>"],                                
                   1.0*inch, 4.9*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    REP_AL.report_context(canvas_rpt, ["<b><i><font color=\"#000080\">"+"No of patients with any specimen +ve for HO "+ACC.dict_org[sh_org][2]+"* : " + str(num_pat_all)+"</font></i></b>"],            
                   1.0*inch, 4.7*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    REP_AL.report_context(canvas_rpt, ["<b><i><font color=\"#000080\">"+"No of wards having patients with any specimens culture +ve for HO "+ACC.dict_org[sh_org][2]+" : " + str(num_ward_all)+"</font></i></b>"], 
                   1.0*inch, 4.5*inch, 460, 50, font_size=11, font_align=TA_LEFT)
    try:
        canvas_rpt.drawImage(filename_graph+"_"+sh_org+"_all.png", 0.1*inch, 2.5*inch, preserveAspectRatio=True, width=8.0*inch, height=2.40*inch,showBoundary=False) 
    except Exception as e:
        #print (e)
        AL.printlog("Error : checkpoint Annex C generate report (graph all) of " + str(sh_org) + " : " +str(e),True,logger) 
        logger.exception(e)
    if len(df_all)>0:
        try:
            REP_AL.report_context(canvas_rpt, ["Table of the two top clusters with p-value<0.05"], 
                   1.0*inch, 1.9*inch, 460, 50, font_size=11, font_align=TA_LEFT)
            table_draw = REP_AL.annexc_table_main(df_all)
            table_draw.wrapOn(canvas_rpt, 500, 300)
            table_draw.drawOn(canvas_rpt, 1.5*inch, 1.3*inch)
        except Exception as e:
            #print (e)
            AL.printlog("Error : checkpoint Annex C generate report (table alld) of " + str(sh_org) + " : " +str(e),True,logger) 
            logger.exception(e)
    else:
        REP_AL.report_context(canvas_rpt, ["There is no cluster with p-value<0.05."], 
                   1.0*inch, 1.7*inch, 460, 50, font_size=11, font_align=TA_LEFT)

    REP_AL.report_context(canvas_rpt,footnote(), 1.0*inch, 0.30*inch, 460, 70, font_size=9,line_space=12)
    #                 report_todaypage(c,55,30,"Created on: "+today)
    REP_AL.report_todaypage(canvas_rpt,270,30,"AnnexC - Page " + str(page_number))
    canvas_rpt.showPage()
    #c.save()
def main_generatepdf(canvas_rpt,logger,startpage):
    bOK = False
    df_num_all = pd.read_excel(ACC.CONST_PATH_ANNEXC_RESULT+ACC.CONST_FILENAME_COUNT+"_all"+".xlsx",index_col=0)
    df_num_blo = pd.read_excel(ACC.CONST_PATH_ANNEXC_RESULT+ACC.CONST_FILENAME_COUNT+"_blo"+".xlsx",index_col=0)
    page = startpage
    #Print main 
    prapare_mainAnnexC_mainpage(canvas_rpt,logger,filename_pdf=ACC.CONST_PATH_ANNEXC_RESULT+ACC.CONST_FILENAME_MAIN_PDF,page_number=page)
    page = page + 1
    for sh_org in ACC.dict_org.keys():
        df_all = prepare_tabletoplot(filename=ACC.CONST_PATH_ANNEXC_RESULT+ACC.CONST_FILENAME_PCLUSTER+"_"+sh_org+"_all.xlsx",lst_tosort=[ACC.CONST_COL_NEWOBS], ascending=False)
        df_blo = prepare_tabletoplot(filename=ACC.CONST_PATH_ANNEXC_RESULT+ACC.CONST_FILENAME_PCLUSTER+"_"+sh_org+"_blo.xlsx",lst_tosort=[ACC.CONST_COL_NEWOBS], ascending=False)
        prapare_mainAnnexC_per_org(canvas_rpt,logger,df_all=df_all, 
                                 df_blo=df_blo, 
                                 sh_org=sh_org,
                                 filename_pdf  =ACC.CONST_PATH_ANNEXC_RESULT+ACC.CONST_FILENAME_MAIN_PDF, 
                                 filename_graph=ACC.CONST_PATH_ANNEXC_RESULT+ACC.CONST_FILENAME_WARDTOP2, 
                                 num_ward_all  =df_num_all.loc[sh_org,ACC.CONST_COL_NUMWARD],
                                 num_ward_blo  =df_num_blo.loc[sh_org,ACC.CONST_COL_NUMWARD], 
                                 num_pat_all   =df_num_all.loc[sh_org,ACC.CONST_COL_NUMOBS], 
                                 num_pat_blo   =df_num_blo.loc[sh_org,ACC.CONST_COL_NUMOBS],
                                 page_number   =page)
        page += 1
    bOK = True
    return bOK
#Call by AMASS core to run generate annex C
def generate_annex_c_report(canvas_rpt,logger,startpage):
    AL.printlog("AMR surveillance report - checkpoint Annex C main page",False,logger)
    if main_preparedata(logger) == True:
        #Calculate number of page for all organism - THIS MAY MOVE INSIDE GENERATE FUNCTION
        #Generate Annex C for all organism
        AL.printlog("AMR surveillance report - checkpoint Annex C for each organism",False,logger)
        if main_generatepdf(canvas_rpt,logger,startpage) == True:    
            #Write success log
            pass
        else:
            AL.printlog("Error : checkpoint Annex C for each organism",False,logger)    
    else:
        AL.printlog("Error : checkpoint Annex C main page",False,logger) 