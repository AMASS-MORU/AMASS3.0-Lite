
#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** CONST FILE and Configurations                                                                   ***#
#***-------------------------------------------------------------------------------------------------***#
# @author: CHALIDA RAMGSIWUTISAK
# Created on: 01 SEP 2023 
import pandas as pd
import psutil,gc
import datetime #for setting date-time format
import logging #for creating logfile
from reportlab.lib.pagesizes import A4 #for setting PDF size
from reportlab.pdfgen import canvas #for creating PDF page
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.lib.styles import ParagraphStyle #for setting paragraph style
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY #for setting paragraph style
from reportlab.platypus import * #for plotting graph and tables
from reportlab.lib.colors import * #for importing color palette
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables
import AMASS_amr_const as AC
import AMASS_amr_const_annex_c as ACC
import AMASS_amr_commonlib as AL
import AMASS_amr_report_new as AMR_REPORT_NEW

#######################################################################
#######################################################################
#####Step1. Merged Hospital-Microbiology data to SaTScan's output######
#######################################################################
#######################################################################
def prepare_fromHospMicro_toSaTScan(logger,df_all=pd.DataFrame(),df_blo=pd.DataFrame()):
    AL.printlog("Start Cluster signal identification (ANNEX C): " + str(datetime.datetime.now()),False,logger)
    for lo_org in ACC.dict_ast.keys():
    # for lo_org in ["organism_staphylococcus_aureus"]:
        lst_usedcolumns = []
        if lo_org in ACC.dict_configuration_astforprofile.keys():
            lst_usedcolumns = [AC.CONST_NEWVARNAME_ORG3]+AC.get_dict_orgcatwithatb(bisabom=True,bisentspp=True)[lo_org][4]+ACC.dict_configuration_astforprofile[lo_org]
        else:
            lst_usedcolumns = [AC.CONST_NEWVARNAME_ORG3]+AC.get_dict_orgcatwithatb(bisabom=True,bisentspp=True)[lo_org][4]
        
        #Blood
        df_blo_str = df_blo.copy()
        #setting values of these columns from categories to string or int
        df_blo_str = assign_strtypetocolumns(df=df_blo_str, lst_col=lst_usedcolumns)
        df_blo_str[[AC.CONST_NEWVARNAME_COHO_FINAL]] = df_blo_str[[AC.CONST_NEWVARNAME_COHO_FINAL]].astype(int)
        #selecting organisms
        df_blo_org = select_dfbyOrganism(logger,df=df_blo_str, col_org=AC.CONST_NEWVARNAME_ORG3, str_selectorg=lo_org)
        #selecting hospital-origin
        df_blo_org_ho = select_ho(logger,df=df_blo_org, col_origin=AC.CONST_NEWVARNAME_COHO_FINAL, num_ho=1)

        #All specimens
        df_all_str = df_all.copy()
        #setting values of these columns from categories to string or int
        df_all_str = assign_strtypetocolumns(df=df_all_str, lst_col=lst_usedcolumns)
        df_all_str[[AC.CONST_NEWVARNAME_COHO_FINAL]] = df_all_str[[AC.CONST_NEWVARNAME_COHO_FINAL]].astype(int)
        #selecting organisms
        df_all_org = select_dfbyOrganism(logger,df=df_all_str, col_org=AC.CONST_NEWVARNAME_ORG3, str_selectorg=lo_org)
        #selecting hospital-origin
        df_all_org_ho = select_ho(logger,df=df_all_org, col_origin=AC.CONST_NEWVARNAME_COHO_FINAL, num_ho=1)

        lst_ris_rpt2=[]
        lst_ris_rpt2_export=[]
        try:
            #getting list of antibiotics for each pathogen i.e. S. aureus 
            lst_ris_rpt2 = get_lstastforpathogen(lo_org=lo_org,check_writereport=False)
            lst_ris_rpt2_export = get_lstastforpathogen(lo_org=lo_org,check_writereport=True)
        except Exception as e:
            AL.printlog("Error, ANNEX C antibiotic selection for profiling: " +  str(e),True,logger)
        for val in ACC.dict_ast[lo_org]:
            atb    = val[0]
            atb_val= val[1]
            sh_org = val[2]
            print ("pathogen:"+sh_org)
            #deduplicating data
            df_dedup_blo = deduplicate(logger, lo_org=lo_org, df=df_blo_org_ho, atb=atb, atb_val=atb_val)
            df_dedup_all = deduplicate(logger, lo_org=lo_org, df=df_all_org_ho, atb=atb, atb_val=atb_val)
            #selecting profiles from configuration
            #There are isolates positive to that pathogen in blood >>> do antibiotics selection for profiling
            lst_ris_profiletemp = []
            lst_ris_profiletemp = select_atbforprofiling(logger, df=df_dedup_blo, lst_col_ris=lst_ris_rpt2)
            print ("Antibiotics for profiling from blood :",[atb.replace("RIS","") for atb in lst_ris_profiletemp])
            #There is no isolate positive to that pathogen in blood >>> do antibiotics selection for profiling by using isolates positive to that pathogen in clinical specimen
            if len(lst_ris_profiletemp)==0:
                lst_ris_profiletemp = select_atbforprofiling(logger, df=df_dedup_all, lst_col_ris=lst_ris_rpt2)
                print ("Antibiotics for profiling from clinical specimens :",[atb.replace("RIS","") for atb in lst_ris_profiletemp])
            else:
                pass

            #profiling
            df_dedup_blo_profile = pd.DataFrame()
            df_dedup_all_profile = pd.DataFrame()
            try:
                df_dedup_blo_profile  = create_risprofile(logger,df=df_dedup_blo, lst_col_ris=lst_ris_rpt2, lst_col_ristemp=lst_ris_profiletemp, col_profile=ACC.CONST_COL_PROFILE, col_profiletemp=ACC.CONST_COL_PROFILETEMP)
                df_dedup_all_profile  = create_risprofile(logger,df=df_dedup_all, lst_col_ris=lst_ris_rpt2, lst_col_ristemp=lst_ris_profiletemp, col_profile=ACC.CONST_COL_PROFILE, col_profiletemp=ACC.CONST_COL_PROFILETEMP)
            except Exception as e:
                AL.printlog("Error, ANNEX C profiling: " +  str(e),True,logger)
            #creating dictionary for profile
            d_profile    = create_dictforMapProfileID_byorg(logger, df_all=df_dedup_all, df_blo=df_dedup_blo, sh_org=sh_org)
            #exporting d_profile to profile_information.xlsx
            export_dprofile(logger, d_profile=d_profile[1], df_groupprofile=d_profile[0], lst_col_ris=lst_ris_rpt2, lst_col_rpt2_export=lst_ris_rpt2_export, filename_profile=ACC.CONST_TEMPDIRPATH+ACC.CONST_FILENAME_PROFILE+"_"+sh_org.upper())
            #mapping profile to dataframe -- blood
            if len(df_dedup_blo_profile)>0:
                df_dedup_blo_profile = map_profileIDtoDataframe(logger, df=df_dedup_blo_profile, d_profile=d_profile[1], sh_org=sh_org, sh_spc="blo", col_profile=ACC.CONST_COL_PROFILE, col_profileid=ACC.CONST_COL_PROFILEID)
                try:
                    df_dedup_blo_profile.to_excel(ACC.CONST_PATIENTIDENPATH+ACC.CONST_FILENAME_HO_DEDUP+"_"+sh_org.upper()+"_"+ACC.dict_spc["blo"]+".xlsx", index=False)
                except Exception as e:
                    AL.printlog("Error, ANNEX C exporting "+ACC.CONST_FILENAME_HO_DEDUP+"_"+sh_org.upper()+"_"+ACC.dict_spc["blo"]+".xlsx"+": " +  str(e),True,logger)
            #mapping profile to dataframe -- all specimens
            if len(df_dedup_all_profile)>0:
                df_dedup_all_profile = map_profileIDtoDataframe(logger, df=df_dedup_all_profile, d_profile=d_profile[1], sh_org=sh_org, sh_spc="all", col_profile=ACC.CONST_COL_PROFILE, col_profileid=ACC.CONST_COL_PROFILEID)
                try:
                    df_dedup_all_profile.to_excel(ACC.CONST_PATIENTIDENPATH+ACC.CONST_FILENAME_HO_DEDUP+"_"+sh_org.upper()+"_"+ACC.dict_spc["all"]+".xlsx", index=False)
                except Exception as e:
                    AL.printlog("Error, ANNEX C exporting "+ACC.CONST_FILENAME_HO_DEDUP+"_"+sh_org.upper()+"_"+ACC.dict_spc["all"]+".xlsx"+": " +  str(e),True,logger)
            print ("-----------------------")
            print ("-----------------------")
            print ("-----------------------")
            #preparing SaTScan's inputs
            evaluation_study = retrieve_startEndDate(filename=ACC.CONST_RESULTDATAPATH+ACC.CONST_FILENAME_REPORT1) #for satscan_param.prm
            for sh_spc in ACC.dict_spc.keys():
                prepare_satscanInput(logger,
                                    filename_isolate =ACC.CONST_FILENAME_HO_DEDUP, 
                                    filename_ward    =ACC.CONST_FILENAME_WARD, 
                                    filename_case    =ACC.CONST_FILENAME_INPUT, 
                                    filename_oriparam=ACC.CONST_FILENAME_ORIPARAM, 
                                    filename_newparam=ACC.CONST_FILENAME_NEWPARAM, 
                                    path_input=ACC.CONST_PATIENTIDENPATH,
                                    path_output=ACC.CONST_TEMPDIRPATH,
                                    sh_org=sh_org, 
                                    sh_spc=sh_spc, 
                                    start_date=evaluation_study[0], 
                                    end_date  =evaluation_study[1])
        del [[df_dedup_blo, df_dedup_blo_profile]]
        del [[df_dedup_all, df_dedup_all_profile]]
        gc.collect()
        df_dedup_blo=pd.DataFrame()
        df_dedup_blo_profile=pd.DataFrame()
        df_dedup_all=pd.DataFrame()
        df_dedup_all_profile=pd.DataFrame()
    del [[df_blo, df_blo_str, df_blo_org, df_blo_org_ho]]
    del [[df_all, df_all_str, df_all_org, df_all_org_ho]]
    gc.collect()
    df_blo=pd.DataFrame()
    df_blo_str=pd.DataFrame()
    df_blo_org=pd.DataFrame()
    df_blo_org_ho=pd.DataFrame()
    df_all=pd.DataFrame()
    df_all_str=pd.DataFrame()
    df_all_org=pd.DataFrame()
    df_all_org_ho=pd.DataFrame()

#####Small functions in the step1.1#####
#Assigning dtype as string for all used columns
#Original dtype is category
def assign_strtypetocolumns(df=pd.DataFrame(), lst_col=[]):
    lst_col_new = []
    for col in lst_col:
        if (col not in lst_col_new) and (col in df.columns):
            lst_col_new.append(col)
    try:
        df[lst_col_new] = df[lst_col_new].astype(str).fillna("").replace("","-")
    except:
        for col in lst_col_new:
            try:
                df[col] = df[col].astype(str).fillna("").replace("","-")
            except:
                pass
    return df
#Parsing df from core AMASS to the step before grouping isolates by specimen date.
#Resistant profile, deduplication
#input : filename of HospMicroData
#output: dataframe of deduplicated hospital-origin isolates with profiles
def deduplicate(logger,lo_org="",df=pd.DataFrame(),atb="",atb_val="",sh_org="", 
                                  col_newdate=AC.CONST_NEWVARNAME_CLEANSPECDATE, 
                                  col_hn     =AC.CONST_NEWVARNAME_HN, col_numamr =AC.CONST_NEWVARNAME_AMR, col_numamrtested =AC.CONST_NEWVARNAME_AMR_TESTED, 
                                  col_profile=ACC.CONST_COL_PROFILE,  col_profiletemp=ACC.CONST_COL_PROFILETEMP, dformat=""):
    df_r_dedup = pd.DataFrame()
    lst_ris_profile = []
    df[atb] = df[atb].astype(str)
    #selecting MRSA, VRE, CREC, CRKP, CRPA, CRAB
    df_r = select_resistantProfile(df=df, d_ast_val=[atb, atb_val, sh_org])
    if len(df_r)>0:
        try:
            #sorting datetime for deduplication
            df_r_dedup = df_r.sort_values(by=[col_hn,col_newdate,col_numamr,col_numamrtested], ascending=[True,True,False,False], na_position="last")
            #deduplication
            df_r_dedup = df_r_dedup.drop_duplicates(col_hn, keep="first")
        except Exception as e:
            AL.printlog("Error, ANNEX C deduplication: " +  str(e),True,logger)
    return df_r_dedup
def summation_numpatient(logger,df=pd.DataFrame(),lst_col=[],col_forgrouping=""):
    df_ = df
    try:
        df_[col_forgrouping] = 1
        df_ = df_.loc[:,lst_col+[col_forgrouping]].groupby([col for col in lst_col if col != col_forgrouping]).sum()
    except Exception as e:
        AL.printlog("Error, ANNEX C counting number of patient: " + str(e),True,logger)
    return df_
#Creating dictionary for mappingID
#ID ordering by all specimen and then blood of each organism
#d_profile["profile_sequence"]="profile_ID"
def create_dictforMapProfileID_byorg(logger, df_all=pd.DataFrame(), df_blo=pd.DataFrame(), sh_org="", col_profileid=ACC.CONST_COL_PROFILEID, col_profile=ACC.CONST_COL_PROFILE, col_profiletemp=ACC.CONST_COL_PROFILETEMP, col_numprofile_all=ACC.CONST_COL_NUMPROFILE_ALL, col_numprofile_blo=ACC.CONST_COL_NUMPROFILE_BLO):
    prefixID=ACC.CONST_PRENAME_PROFILEID+sh_org+"_"
    df_all_num = summation_numpatient(logger,df_all,lst_col=[col_profile,col_profiletemp],col_forgrouping=col_numprofile_all)
    df_blo_num = summation_numpatient(logger,df_blo,lst_col=[col_profile,col_profiletemp],col_forgrouping=col_numprofile_blo)
    df = pd.DataFrame(columns=[col_profiletemp,col_profile,col_numprofile_all,col_numprofile_blo])
    df = pd.concat([df_blo_num,df_all_num],axis=1).fillna(0).astype(int).reset_index()
    df[col_profileid] = prefixID+"na"
    count = 0
    d_profile = {}
    for idx in df.index:
        try:
            if df.loc[idx,col_profiletemp] == "":
                df.at[idx,col_profileid] = prefixID+"na"
            else:
                count += 1
                df.at[idx,col_profileid] = prefixID+str(count)
            d_profile[df.loc[idx,col_profile]] = df.loc[idx,col_profileid]
        except Exception as e:
            AL.printlog("Error, ANNEX C assigning profile ID: " + str(e),True,logger)
    df = df.loc[:,[col_profileid,col_profile,col_numprofile_blo,col_numprofile_all]]
    return df, d_profile

def map_profileIDtoDataframe(logger, df=pd.DataFrame(), d_profile={}, filename_dedup="", sh_org="", sh_spc="", 
                             col_profile=ACC.CONST_COL_PROFILE, col_profileid=ACC.CONST_COL_PROFILEID):
    try:
        df[col_profileid] = df[col_profile].map(d_profile).fillna("")
        return df
    except Exception as e:
        AL.printlog("Error, ANNEX C mapping profile id: " + str(e),True,logger)
#Creating profile information.xlsx
#column profile_ID : profile_CRAB_1, profile_CRAB_2, profile_CRAB_3, ..., profile_XXXX_N
#column RISdrug1 : -, -, R, S
#column RISdrug2 : R, -, -, -
#column RISdrugN : -, -, R, S
def export_dprofile(logger, d_profile={}, df_groupprofile=pd.DataFrame(), lst_col_ris=[], lst_col_rpt2_export=[], filename_profile="", 
                    col_profileid=ACC.CONST_COL_PROFILEID, col_profile=ACC.CONST_COL_PROFILE, col_numprofile_all=ACC.CONST_COL_NUMPROFILE_ALL, col_numprofile_blo=ACC.CONST_COL_NUMPROFILE_BLO):
    if len(df_groupprofile) > 0:
        try:
            df_profile = df_groupprofile.loc[:,[col_profileid,col_profile]]
            df_profile[lst_col_ris] = df_groupprofile[col_profile].str.split(";",expand=True)
            df_profile = df_profile.merge(df_groupprofile.loc[:,[col_profileid,col_profile,col_numprofile_blo,col_numprofile_all]],left_on=col_profile, right_on=col_profile).rename(columns={col_profileid+"_x":col_profileid}).drop(columns=[col_profileid+"_y",col_profile])
            df_profile.columns = [col_profileid]+lst_col_rpt2_export+[col_numprofile_blo,col_numprofile_all]
            df_profile.to_excel(filename_profile+".xlsx", index=False)
        except Exception as e:
            AL.printlog("Error, ANNEX C exporting profile_information.xlsx: " + str(e),True,logger)
    else:
        pass
#selecting organisms
def select_dfbyOrganism(logger, df=pd.DataFrame(), col_org="", str_selectorg=""):
    df_new = pd.DataFrame()
    try:
        df_new = df.loc[(df[col_org]==str_selectorg)]
    except Exception as e:
        AL.printlog("Error, ANNEX C Selecting organism: " +  str(e),True,logger)
    return df_new
#hospitalized selection
def select_ho(logger, df=pd.DataFrame(), col_origin="", num_ho=0):
    df_ho = pd.DataFrame()
    try:
        df_ho = df.loc[(df[col_origin]==num_ho)]
    except Exception as e:
        AL.printlog("Error, ANNEX C Hospital-origin isolate selection: " +  str(e),True,logger)
    return df_ho
#resistant isolate selection
def select_resistantProfile(df=pd.DataFrame(), d_ast_val=[]):
    df_ast = pd.DataFrame()
    if len(d_ast_val) == 3:
        df_ast = df.loc[(df[d_ast_val[0]]==str(d_ast_val[1]))]
    elif len(d_ast_val) > 3:
        i = 0
        df_ast = df.copy()
        while (i < len(d_ast_val)) and (i+1!=len(d_ast_val)): #in the case; ["RIS3gc","R","RISCarbapenem","S","3gcr-csec"] >>> do not retrieve "3gcr-crec"
            df_ast = df_ast.loc[(df[d_ast_val[i]]==str(d_ast_val[i+1]))]
            i+=2
    return df_ast
#Getting list of antibiotics from report2 i.e. S. aureus is ['RISVancomycin', 'RISClindamycin', 'RISCefoxitin', 'RISOxacillin']
def get_lstastforpathogen(lo_org="",check_writereport=False):
    lst_atb      = []
    lst_atbgroup = []
    #getting list ['RIS3gc','RISCarbapenem','RISFluoroquin','RISTetra','RISaminogly','RISmrsa','RISpengroup']
    for val in AC.dict_atbgroup().values():
        lst_atbgroup.append(val[0])
    #getting list of antibiotics for analysis(column name) or report
    d_ = AC.get_dict_orgcatwithatb(bisabom=True,bisentspp=True)
    v = [i for i in range(len(d_[lo_org][4])) if d_[lo_org][4][i] not in lst_atbgroup]
    if check_writereport:
        lst_atb = [d_[lo_org][3][idx].replace("RIS","") for idx in v ]
    else:
        lst_atb = [d_[lo_org][4][idx] for idx in v ]
    #getting list of additional antibiotics from configuration
    if (ACC.dict_configuration_astforprofile != {}) and (lo_org in ACC.dict_configuration_astforprofile.keys()):
        lst_atb = lst_atb + ACC.dict_configuration_astforprofile[lo_org]
        lst_atb_unique = []
        for atb in lst_atb:
            if atb not in lst_atb_unique:
                if check_writereport:
                    lst_atb_unique.append(atb.replace("RIS",""))
                else:
                    lst_atb_unique.append(atb)
            else:
                pass
    else:
        lst_atb_unique = lst_atb
    return lst_atb_unique
#Selecting range of criteria for selecting antibiotics profiling
#In the case that there is no setting >>> using default setting
def select_rateRangeforprofiling(configuration_profile=ACC.dict_configuration_profile,param=""):
    try:
        min_ = float(configuration_profile[param][0])
    except:
        if ACC.CONST_VALUE_TESTATBRATE == param:
            min_ = 91
        else:
            min_ = 0.1
    try:
        max_ = float(configuration_profile[param][1])
    except:
        if ACC.CONST_VALUE_TESTATBRATE == param:
            min_ = 100
        else:
            max_ = 99.9
    return min_,max_
#Counting number of resistant, intermediate, and susceptible
#In the case that there is no setting >>> assign 0
def count_numRISbyATB(df=pd.DataFrame(),col_atb=[],val_forcount=""):
    num = 0
    try:
        num = int(df[col_atb].str.count(val_forcount).sum())
    except: #no val_forcount in a column >>> num = 0
        pass
    return num
#Selecting list of ATBs which pass the criteria for profiling
def select_atbforprofiling(logger,df=pd.DataFrame(), lst_col_ris=[], configuration_profile=ACC.dict_configuration_profile,
                          resistant=ACC.dict_ris["resistant"],intermediate=ACC.dict_ris["intermediate"],susceptible=ACC.dict_ris["susceptible"]):
    lst_col_ris_include = []
    for atb in lst_col_ris:
        try:
            if atb in df.columns:
                num_r = count_numRISbyATB(df=df,col_atb=atb,val_forcount=resistant)
                num_i = count_numRISbyATB(df=df,col_atb=atb,val_forcount=intermediate)
                num_s = count_numRISbyATB(df=df,col_atb=atb,val_forcount=susceptible)
                total_testedatb = num_r+num_i+num_s
                min_testedatb = select_rateRangeforprofiling(param=ACC.CONST_VALUE_TESTATBRATE)[0]
                max_testedatb = select_rateRangeforprofiling(param=ACC.CONST_VALUE_TESTATBRATE)[1]
                if (total_testedatb*100/len(df)>=min_testedatb) and (total_testedatb*100/len(df)<=max_testedatb):
                    for ris in [keys for keys in ACC.dict_configuration_profile.keys() if keys != ACC.CONST_VALUE_TESTATBRATE]:
                        min_testedris = select_rateRangeforprofiling(param=ris)[0]
                        max_testedris = select_rateRangeforprofiling(param=ris)[1]
                        numerator = 0
                        if ACC.CONST_VALUE_RRATE == ris:
                            numerator = num_r
                        elif ACC.CONST_VALUE_IRATE == ris:
                            numerator = num_i
                        elif ACC.CONST_VALUE_SRATE == ris:
                            numerator = num_s
                        #calculating rate
                        if (numerator*100/total_testedatb>=min_testedris) and (numerator*100/total_testedatb<=max_testedris):
                            print (atb, "%tested AST:"+str(round(total_testedatb*100/len(df),ndigits=2)), "%R:"+str(round(num_r*100/total_testedatb,ndigits=2)), "%S:"+str(round(num_s*100/total_testedatb,ndigits=2)))
                            lst_col_ris_include.append(atb)
                            break
        except Exception as e:
            AL.printlog("Error, ANNEX C selecting Antibiotics for profiling: " +  str(e),True,logger)
    lst_col_ris_include_unique = []
    for atb in lst_col_ris_include:
        if atb not in lst_col_ris_include_unique:
            lst_col_ris_include_unique.append(atb)
    return lst_col_ris_include_unique
#Creating "R---I-S---R..."
#lst_col_ris is list of full antibiotics recommended use from CLSI for that pathogen
#lst_col_ristemp is list of antibiotics for profiling
def create_risprofile(logger,df=pd.DataFrame(), lst_col_ris=[], lst_col_ristemp=[], col_profile="", col_profiletemp=""):
    df[col_profile] = ""
    df[col_profiletemp] = ""
    for atb in lst_col_ris:
        if atb in lst_col_ristemp:
            try:
                df[col_profile] = df[col_profile] + ";" + df[atb].astype(str)
                df[col_profiletemp] = df[col_profiletemp] + ";" + df[atb].astype(str)
            except Exception as e:
                AL.printlog("Error, ANNEX C creating profile column: " + str(e),True,logger)
        else:
            #assign to pos of antibiotic that DO NOT INCLUDED (NC) to profiling >>> need to rename from "NC" to "-" before returning df
            df[col_profile] = df[col_profile] + ";" + "NC"
    # - >>> DO NOT DONE (ND) or not tested but included for profiling
    #DO NOT INCLUDED (NC) to profiling >>> -
    df[col_profile] = df[col_profile].str.replace("-","ND").str.replace("NC","-").str[1:]
    df[col_profiletemp] = df[col_profiletemp].str.replace("-","ND").str[1:]
    return df
#Droping RIS duplicated sequences
def drop_dupProfile(df=pd.DataFrame(), col_profile=""):
    return df[col_profile].drop_duplicates(keep="first")
def create_dictforMapProfileID(lst_profile=[], prefixID=""):
    d_profile = {}
    for i in range(len(lst_profile)):
        d_profile[lst_profile[i]] = prefixID+str(i+1) #profile should start from 1
    return d_profile

#####Small functions in the step1.2#####
#Parsing df from the step before grouping isolates by specimen date to satscan_input.csv, satscan_location.csv, and satscan_param.prm.
#Grouping isolates by specimen collection date, prepare location and prm files
#input : excel file of dataframe of deduplicated hospital-origin isolates with profiles
#output: csv files for satscan_input.csv, satscan_location.csv, and satscan_param.prm
def prepare_satscanInput(logger,filename_isolate="", filename_ward="", filename_case="", filename_oriparam="", filename_newparam="", path_input="", path_output="",
                         sh_org="", sh_spc="", start_date="",  end_date="", str_ward=AC.CONST_VARNAME_WARD, 
                         col_amass=ACC.CONST_COL_AMASSNAME,    col_user=ACC.CONST_COL_USERNAME,        col_resist=ACC.CONST_COL_RESISTPROFILE,
                         col_mrsa=AC.CONST_NEWVARNAME_ASTMRSA_RIS,col_3gc=AC.CONST_NEWVARNAME_AST3GC_RIS,    col_crab=AC.CONST_NEWVARNAME_ASTCBPN_RIS,  col_van=ACC.CONST_NEWVARNAME_ASTVAN_RIS,
                         col_org=AC.CONST_NEWVARNAME_ORG3,    col_oriorg=AC.CONST_VARNAME_ORG,       col_profile=ACC.CONST_COL_PROFILEID, 
                         col_wardname=AC.CONST_VARNAME_WARD ,  col_wardid=AC.CONST_NEWVARNAME_WARDCODE,col_cleandate=AC.CONST_NEWVARNAME_CLEANSPECDATE, 
                         col_testgroup=ACC.CONST_COL_TESTGROUP,col_numcase=ACC.CONST_COL_CASE):
    df = pd.DataFrame()
    directory_path = path_input
    file_name = filename_isolate+"_"+sh_org.upper()+"_"+ACC.dict_spc[sh_spc]+".xlsx"
    file_path = os.path.join(directory_path, file_name)

    if os.path.exists(file_path):
        #For satscan_input.csv
        try:
            df = pd.read_excel(path_input+filename_isolate+"_"+sh_org.upper()+"_"+ACC.dict_spc[sh_spc]+".xlsx").fillna("")
            df = df.loc[:,[col_wardname,col_wardid,col_oriorg,col_org,col_cleandate,col_profile,col_mrsa,col_3gc,col_crab,col_van]]
        except Exception as e:
            AL.printlog("Error, ANNEX C reading deduplicated microbiology file: " +  str(e),True,logger)
        df_ns = pd.DataFrame(columns=[col_testgroup,col_cleandate,col_numcase])
        try:
            d_org = ACC.dict_org[sh_org][0] #get full org name i.e. 'organism_acinetobacter_baumannii'
            for value in ACC.dict_ast[d_org]:
                if sh_org in value:
                    df_temp = df.copy()
                    df_temp[col_resist] = value[-1] #put short resistant name i.e. crab
                    ##Grouping records by cluster_code and specimen_collection_date
                    df_temp = create_clusterCode(df=df_temp, col_testgroup=col_testgroup, col_ward=col_wardid, col_org=col_resist, col_profile=col_profile)
                    df_temp[col_numcase] = 1
                    df_temp_code = df_temp.groupby(by=[col_testgroup,col_cleandate]).sum().reset_index()
                    df_ns = pd.concat([df_ns,df_temp_code],axis=0)
                    df_ns = df_ns.reset_index().loc[:,[col_testgroup,col_numcase,col_cleandate]]
                    df_ns = df_ns.sort_values(by=[col_cleandate]).reset_index().drop(columns=["index"])
                    try:
                        df_ns.to_csv(path_output+filename_case+"_"+sh_org+"_"+sh_spc+".csv",encoding="ascii",index=False)
                    except Exception as e:
                        AL.printlog("Error, ANNEX C SaTScan input exportation: " +  str(e),True,logger)
                else:
                    pass
        except Exception as e:
            AL.printlog("Error, ANNEX C Case file preparation: " +  str(e),True,logger)
        #For satscan_location.csv
        try:
            prepare_locfile(lst_clustercode=df_ns[col_testgroup].unique(), df_ward_ori=df.loc[:,[col_wardid,col_wardname]], col_wardid=col_wardid, col_testgroup=col_testgroup,  col_wardname=col_wardname, col_profileid=col_profile, col_resistprofile=col_resist, sh_org=sh_org, sh_spc=sh_spc, path_output=path_output)
        except Exception as e:
            AL.printlog("Error, ANNEX C Location file preparation: " +  str(e),True,logger)
        #For satscan_param.prm
        try:
            prepare_prmfile(ori_prmfile=filename_oriparam, new_prmfile=filename_newparam, 
                            start_date=start_date, end_date=end_date, path_output=path_output,sh_org=sh_org, sh_spc=sh_spc)
        except Exception as e:
            AL.printlog("Error, ANNEX C Parameter file preparation: " +  str(e),True,logger)
    else:
        pass

    
def read_dictionary(prefix_filename=""):
    df = pd.DataFrame()
    try:
        df = pd.read_excel(prefix_filename+".xlsx")
    except:
        try:
            df = pd.read_csv(prefix_filename+".csv")
        except:
            df = pd.read_csv(prefix_filename+".csv", encoding="windows-1252")
    return df.iloc[:,:2].fillna("")
#Retrieving user value from dictionary_for_microbiology_data.xlsx
def retrieve_uservalue(df, amass_name, col_amass="", col_user=""):
    return df.loc[df[col_amass]==amass_name,:].reset_index().loc[0,col_user]
#Retrieve date from Report1.csv
def retrieve_startEndDate(filename="",col_datafile=ACC.CONST_COL_DATAFILE,col_param=ACC.CONST_COL_PARAM,val_datafile=ACC.CONST_VALUE_DATAFILE,val_sdate=ACC.CONST_VALUE_SDATE,val_edate=ACC.CONST_VALUE_EDATE,col_date=ACC.CONST_COL_DATE):
    df = pd.read_csv(filename)
    start_date = get_dateforprm(df.loc[(df[col_datafile]==val_datafile)&(df[col_param]==val_sdate),col_date].tolist()[0])
    end_date   = get_dateforprm(df.loc[(df[col_datafile]==val_datafile)&(df[col_param]==val_edate),col_date].tolist()[0])
    return start_date, end_date
#Preparing date from 01 Dec 2022 to 2022/12/01
def get_dateforprm(date=""):
    fmt_date = ""
    month = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05","Jun":"06",
             "Jul":"07","Aug":"08","Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
    for keys,values in month.items():
        if keys in date:
            map_date = date.replace(keys,values).split(" ")
            fmt_date = str(map_date[2])+"/"+str(map_date[1])+"/"+str(map_date[0])
            break
    return fmt_date
#Finding a dictionary key from value
def get_keysfromvalues(str_finder="", d_={}):
    for keys,values in d_.items():
        if values == str_finder:
            return keys
#create a cluster code column
def create_clusterCode(df=pd.DataFrame(), col_testgroup="", col_ward="", col_org="", col_profile=""):
    df[col_testgroup] = df[col_ward] + ";" + df[col_profile] + ";" + df[col_org]
    return df
#Prepare parameter file
def prepare_prmfile(ori_prmfile="",start_date="", end_date="", new_prmfile="", sh_org="", sh_spc="", path_output=""):
    str_wholefile = ""
    file = open(ori_prmfile,"r", encoding="ascii")
    f = file.readline()
    while f != "":  
        for keys,values in ACC.dict_configuration_prm.items():
            if keys in f:
                if ("CaseFile=" == keys) or ("CoordinatesFile=" == keys) or ("ResultsFile=" == keys):
                    f = keys+"./"+values+sh_org+"_"+sh_spc+".csv" + "\n"
                    break
                else:
                    f = keys + values + "\n"
                    break
            else:
                if ("StartDate=" in f) and ("ProspectiveStartDate=" not in f):
                    f = "StartDate="+ start_date + "\n"
                    break
                elif ("EndDate=" in f):
                    f = "EndDate="+ end_date + "\n"
                    break
                else:
                    pass
        str_wholefile = str_wholefile + f
        f = file.readline()
    file.close()
    filename_prm = open(path_output+new_prmfile+"_"+sh_org+"_"+sh_spc+".prm", "w")
    filename_prm.write("".join(str_wholefile))
    filename_prm.close()
#Prepare location file
def prepare_locfile(lst_clustercode=[], df_ward_ori=pd.DataFrame(),
                    col_testgroup="", col_wardid="", col_wardname="", col_profileid="", col_resistprofile="", 
                    sh_org="", sh_spc="", path_output=""):
    df = pd.DataFrame(lst_clustercode,columns=[col_testgroup]).reset_index().drop(columns=["index"])
    df[["x_axis","y_axis"]]=0
    for idx in df.index:
        df.at[idx,"x_axis"]=idx*3
        df.at[idx,"y_axis"]=idx*3
    df[col_testgroup+"_imd"]=df[col_testgroup].str.split(";")
    df[col_wardid]            =df[col_testgroup+"_imd"].str[0]
    df[col_profileid]         =df[col_testgroup+"_imd"].str[1]
    df[col_resistprofile]     =df[col_testgroup+"_imd"].str[2]
    df = df.drop(columns=[col_testgroup+"_imd"])
    df_ward_ori = df_ward_ori.drop_duplicates(keep="first")
    df_2 = df_ward_ori.merge(df, left_on=col_wardid, right_on=col_wardid,how="inner")
    df_2.loc[:,[col_testgroup,"x_axis","y_axis"]].to_csv(path_output+ACC.CONST_FILENAME_LOCATION+"_"+sh_org+"_"+sh_spc+".csv",encoding="ascii",errors="replace",index=False)

#######################################################################
#######################################################################
################# Step2. Performing SaTScan program ###################
#######################################################################
#######################################################################
import os
import subprocess
def call_SaTScan(prmfile=""):
    for sh_org in ACC.dict_org.keys():
        for sh_spc in ACC.dict_spc.keys():
            print ("Performing cluster detection for "+ str(sh_org).upper() + " in " + str(ACC.dict_spc[sh_spc]))
            subprocess.run(".\Programs\AMASS_amr\SaTScanBatch64.exe -f "+ prmfile+"_"+sh_org+"_"+sh_spc+".prm", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
#######################################################################
#######################################################################
################## Step3. Visualization for AnnexC ####################
#######################################################################
#######################################################################
def prepare_annexc_results(logger,b_wardhighpat=True, num_wardhighpat=2):
    for sh_spc in ACC.dict_spc.keys():
        df_num = pd.DataFrame(0,index=ACC.dict_org.keys(),columns=[ACC.CONST_COL_CASE,ACC.CONST_COL_WARDID])
        for sh_org in ACC.dict_org.keys():
            #read satscan result
            df_re = pd.DataFrame()
            try:
                df_re = read_texttocol(ACC.CONST_TEMPDIRPATH+ACC.CONST_FILENAME_RESULT+"_"+sh_org+"_"+sh_spc+".col.txt")
                #ward_015;Profile_7;crab;ward_model
                df_re[[ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_RESISTPROFILE]] = df_re.loc[:,ACC.CONST_COL_LOCID].str.split(";", expand=True)
            except Exception as e:
                df_re[[ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_RESISTPROFILE]] = ""

            #formatting datetime
            if len(df_re) > 0:
                df_re_date = AL.fn_clean_date(df=df_re,     oldfield=ACC.CONST_COL_SDATE,cleanfield=ACC.CONST_COL_CLEANSDATE,dformat="",logger="")
                df_re_date = AL.fn_clean_date(df=df_re_date,oldfield=ACC.CONST_COL_EDATE,cleanfield=ACC.CONST_COL_CLEANEDATE,dformat="",logger="")
                df_re_date[ACC.CONST_COL_CLEANSDATE] = df_re_date[ACC.CONST_COL_CLEANSDATE].astype(str)
                df_re_date[ACC.CONST_COL_CLEANEDATE] = df_re_date[ACC.CONST_COL_CLEANEDATE].astype(str)
                df_re_date_pval = format_lancent_pvalue(df=df_re_date,
                                                        col_pval=ACC.CONST_COL_PVAL,
                                                        col_clean_pval=ACC.CONST_COL_CLEANPVAL)
            else:
                df_re_date = df_re.copy()
                df_re_date[[ACC.CONST_COL_CLEANPVAL,ACC.CONST_COL_CLEANSDATE,ACC.CONST_COL_CLEANEDATE]] = ""

            #preparing date and export results
            select_col = [AC.CONST_VARNAME_ORG,ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_CLEANSDATE,ACC.CONST_COL_CLEANEDATE,
                          ACC.CONST_COL_OBS,ACC.CONST_COL_CLEANPVAL,ACC.CONST_COL_CLEANPVAL+"_lancent"]
            df_re_date_1 = pd.DataFrame(columns=select_col)
            if len(df_re_date) > 0:
                try:
                    df_re_date_1 = df_re_date.copy()
                    df_re_date_1[AC.CONST_VARNAME_ORG]   = df_re_date_1[ACC.CONST_COL_RESISTPROFILE].map(ACC.dict_org)
                    df_re_date_1[[AC.CONST_VARNAME_ORG]] = df_re_date_1[AC.CONST_VARNAME_ORG][0][1]
                    df_re_date_1 = df_re_date_1.loc[:,select_col]
                    df_re_date_2 = df_re_date_1.loc[df_re_date_1[ACC.CONST_COL_CLEANPVAL].astype(float)<0.05,:]
                    reorder_clusters(df=pd.concat([df_re_date_2,df_re_date_1.loc[df_re_date_1[ACC.CONST_COL_CLEANPVAL].astype(float)>=0.05,:]])).rename(columns={ACC.CONST_COL_WARDID:ACC.CONST_COL_WARDID,
                                                                                                                                                                ACC.CONST_COL_CLEANSDATE:ACC.CONST_COL_NEWSDATE,
                                                                                                                                                                ACC.CONST_COL_CLEANEDATE:ACC.CONST_COL_NEWEDATE,
                                                                                                                                                                ACC.CONST_COL_OBS:ACC.CONST_COL_NEWOBS, 
                                                                                                                                                                ACC.CONST_COL_CLEANPVAL:ACC.CONST_COL_NEWPVAL,
                                                                                                                                                                ACC.CONST_COL_CLEANPVAL+"_lancent":ACC.CONST_COL_NEWPVAL+"lancent"}).to_excel(ACC.CONST_TEMPDIRPATH+ACC.CONST_FILENAME_ACLUSTER+"_"+str(sh_org).upper()+"_"+str(sh_spc)+".xlsx",index=False,header=True)
                    retrieve_databypval(df=reorder_clusters(df=df_re_date_1),col_pval=ACC.CONST_COL_CLEANPVAL).rename(columns={ACC.CONST_COL_WARDID:ACC.CONST_COL_WARDID,
                                                                                                                                ACC.CONST_COL_CLEANSDATE:ACC.CONST_COL_NEWSDATE,
                                                                                                                                ACC.CONST_COL_CLEANEDATE:ACC.CONST_COL_NEWEDATE,
                                                                                                                                ACC.CONST_COL_OBS:ACC.CONST_COL_NEWOBS, 
                                                                                                                                ACC.CONST_COL_CLEANPVAL+"_lancent":ACC.CONST_COL_NEWPVAL}).drop(columns=[ACC.CONST_COL_CLEANPVAL]).to_excel(ACC.CONST_TEMPDIRPATH+ACC.CONST_FILENAME_PCLUSTER+"_"+str(sh_org).upper()+"_"+str(sh_spc)+".xlsx",index=False,header=True)
                except Exception as e:
                    df_re_date_1.drop(columns=[ACC.CONST_COL_CLEANPVAL+"_temp"])
                    AL.printlog("Error, ANNEX C "+ACC.CONST_FILENAME_PCLUSTER+"_"+str(sh_org)+"_"+str(sh_spc)+".xlsx: " + str(e),True,logger)
            else:
                pass

#####Small functions in the step3#####
def create_df_forstartweekday(s_studydate="2021/01/01", e_studydate="2021/12/31", fmt_studydate="%Y/%m/%d",
                              col_sweekday="startweekday"):
    from datetime import date, timedelta, datetime
    s_studyweek = int(datetime.strptime(s_studydate,fmt_studydate).strftime("%W"))
    s_studyyear = datetime.strptime(s_studydate,fmt_studydate).year
    e_studyyear = datetime.strptime(e_studydate,fmt_studydate).year
    #if study period is not in the first week of that year >>> set the first week
    if s_studyweek != 0:
        s_studyweek = int(datetime.strptime(str(s_studydate),fmt_studydate).strftime("%W"))
    else:
        pass
    #whether study period is over one year >>> include every years during study period
    if e_studyyear > s_studyyear:
        i = int(s_studyyear)
        e_studyweek = int(datetime.strptime(str(e_studydate),fmt_studydate).strftime("%W"))
        while i < e_studyyear:
            e_studyweek = e_studyweek+int(datetime.strptime(str(i)+"/12/31",fmt_studydate).strftime("%W"))
            i += 1
            print ("year: " + str(i))
    else:
        e_studyweek = int(datetime.strptime(str(e_studyyear)+"/12/31",fmt_studydate).strftime("%W"))
    total_studyday = (datetime.strptime(str(e_studyyear)+"/12/31",fmt_studydate) - datetime.strptime(str(s_studyyear)+"/01/01",fmt_studydate)).days+1

    lst_startweekday = get_startweekday_forXyear(s_studydate=str(s_studyyear)+"/01/01",
                                                 fmt_studydate  =fmt_studydate,
                                                 total_studyday =total_studyday,
                                                 total_studyweek=e_studyweek-s_studyweek+1)
    df_graph = pd.DataFrame("",
                        index=range(0,int(e_studyweek+1)), 
                        columns=[col_sweekday])
    for idx in df_graph.index:
        df_graph.at[idx,col_sweekday] = str(idx) +" (" + str(lst_startweekday[idx]) + ")"
    return df_graph
def ori_create_df_forstartweekday(s_studydate="2021/01/01", e_studydate="2021/12/31", fmt_studydate="%Y/%m/%d",
                              col_sweekday="startweekday"):
    from datetime import date, timedelta, datetime
    s_studyweek = int(datetime.strptime(s_studydate,fmt_studydate).strftime("%W"))
    s_studyyear = datetime.strptime(s_studydate,fmt_studydate).year
    e_studyyear = datetime.strptime(e_studydate,fmt_studydate).year
    #if study period is not in the first week of that year >>> set the first week
    if s_studyweek != 0:
        s_studyweek = int(datetime.strptime(str(s_studyyear)+"/01/01",fmt_studydate).strftime("%W"))
    else:
        pass
    #whether study period is over one year >>> include every years during study period
    if e_studyyear > s_studyyear:
        i = int(s_studyyear)
        e_studyweek = int(datetime.strptime(str(s_studyyear)+"/12/31",fmt_studydate).strftime("%W"))
        while i < e_studyyear:
            e_studyweek = e_studyweek+int(datetime.strptime(str(i)+"/12/31",fmt_studydate).strftime("%W"))
            i += 1
            print ("year: " + str(i))
    else:
        e_studyweek = int(datetime.strptime(str(e_studyyear)+"/12/31",fmt_studydate).strftime("%W"))
    total_studyday = (datetime.strptime(str(e_studyyear)+"/12/31",fmt_studydate) - datetime.strptime(str(s_studyyear)+"/01/01",fmt_studydate)).days+1

    lst_startweekday = get_startweekday_forXyear(s_studydate=str(s_studyyear)+"/01/01",
                                                 fmt_studydate  =fmt_studydate,
                                                 total_studyday =total_studyday,
                                                 total_studyweek=e_studyweek-s_studyweek+1)
    df_graph = pd.DataFrame("",
                        index=range(0,int(e_studyweek+1)), 
                        columns=[col_sweekday])
    for idx in df_graph.index:
        df_graph.at[idx,col_sweekday] = str(idx) +" (" + str(lst_startweekday[idx]) + ")"
    return df_graph
#Assigning number of patient and number of wards
def assign_valuetodf(logger,df=pd.DataFrame(),idx_assign="",col_assign="",val_toassign=0):
    try:
        df.at[idx_assign,col_assign] = val_toassign
    except Exception as e:
        AL.printlog("Error, ANNEX C assigning number of patients and number of wards to a dataframe: " + str(e),True,logger)
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
def read_texttocol(filename = ""):
    df=pd.DataFrame()
    try:
        file = open(filename, 'r')
        line = file.readline()
        lst_text = []
        while line!="":
            split = [i for i in line.replace("\n","").split(" ") if i != ""]
            lst_text.append(split)
            line = file.readline()
        file.close()
        df = pd.DataFrame(lst_text[1:],columns=lst_text[0]).fillna("NA")
    except:
        pass
    return df
def format_lancent_pvalue(df=pd.DataFrame(),col_pval="",col_clean_pval=""):
    df[col_clean_pval] = ""
    for idx in df.index:
        ori_pval = df.loc[idx,col_pval]
        new_pval_lancent = ""
        new_pval_temp = float(0)
        if float(ori_pval)>=0.995:
            new_pval_temp = round(float(ori_pval),ndigits=2)
            new_pval_lancent = ">0.99"
        elif (float(ori_pval)>0.1) and (float(ori_pval)<0.995):
            new_pval_temp = round(float(ori_pval),ndigits=2)
            new_pval_lancent = str(new_pval_temp)
        elif (float(ori_pval)>0.01) and (float(ori_pval)<=0.1):
            new_pval_temp = round(float(ori_pval),ndigits=3)
            new_pval_lancent = str(new_pval_temp)
        elif (float(ori_pval)>0.001) and (float(ori_pval)<=0.01):
            new_pval_temp = round(float(ori_pval),ndigits=4)
            new_pval_lancent = str(new_pval_temp)
        elif (float(ori_pval)>0.0001) and (float(ori_pval)<=0.001):
            new_pval_temp = round(float(ori_pval),ndigits=5)
            new_pval_lancent = str(new_pval_temp)
        elif (float(ori_pval)>0.00001) and (float(ori_pval)<=0.0001):
            new_pval_temp = round(float(ori_pval),ndigits=6)
            new_pval_lancent = "<0.0001"
        df.at[idx,col_clean_pval+"_lancent"] = new_pval_lancent
        df.at[idx,col_clean_pval] = new_pval_temp
    return df
#retrieving records by pval
def retrieve_databypval(df=pd.DataFrame(),col_pval="",pval=0.05):
    return df.loc[df[col_pval].astype(float)<pval,:]
#backup function for several years of startweekday
def get_startweekday_forXyear(s_studydate="2021/01/01",fmt_studydate="%Y/%m/%d",total_studyday=0,total_studyweek=0):
    from datetime import date, timedelta, datetime
    lst_startweek = []
    today  = datetime.strptime(s_studydate, fmt_studydate).date()
    numday = 0
    start_week0 = today - timedelta(days=today.weekday())
    lst_startweek.append(str(start_week0))
    while numday <= total_studyday:
        numday += 7
        start_weekX = start_week0 + timedelta(days=numday)
        lst_startweek.append(str(start_weekX))
    return lst_startweek[:total_studyweek]
#Plotting AnnexC clusters
def plot(logger, df_org=pd.DataFrame(), palette=[], filename="",
         xlabel="Number of week (Start of weekday)",ylabel="Number of patients(n)*",
         min_yaxis=0,max_yaxis=5,step=2):
    import matplotlib.pyplot as plt
    import numpy as np
    try:
        plt.figure()
        df_org.plot(kind='bar', 
                stacked =True, 
                figsize =(20, 10), 
                color =palette, 
                fontsize=16,edgecolor="black",linewidth=2)
        plt.legend(prop={'size': 14},ncol=4)
        plt.ylabel(ylabel, fontsize=16)
        plt.xlabel(xlabel, fontsize=16)
        plt.yticks(np.arange(min_yaxis,max_yaxis+1,step=step), fontsize=16)
        plt.savefig(filename, format='png',dpi=180,transparent=True, bbox_inches="tight")
        plt.close()
        plt.clf
    except Exception as e:
        AL.printlog("Error, ANNEX C Plotting ANNEXC graph: " + str(e),True,logger)


#Reading and preparing satscan_input.csv
def read_satscaninput(filename="", col_code="", col_wardid="", col_profileid="", col_resist="", col_spcdate="", col_numweek="", col_wardprof=""):
    df_input = pd.DataFrame()
    try:
        df_input = pd.read_csv(filename).reset_index().drop(columns=["index"])
        df_input[[col_wardid,col_profileid,col_resist]] = df_input.loc[:,col_code].str.split(";", expand=True)
        df_input[col_numweek] = pd.to_datetime(df_input[col_spcdate]).dt.strftime("%W").astype(int)
        df_input[col_wardprof] = df_input[col_wardid]+";"+df_input[col_profileid]
    except:
        df_input[[col_code,ACC.CONST_COL_CASE,col_spcdate,
                  col_wardid,col_profileid,col_resist,
                  col_numweek,col_wardprof]] = ""
    return df_input

def assign_numcase(df_graph=pd.DataFrame(),df_data=pd.DataFrame(),lst_topward=[],col_wardid="",col_numweek="",col_numcase=""):
    df_graph_new = df_graph.copy()
    df_graph_new[lst_topward] = 0
    for idx in df_data.index:
        if df_data.loc[idx,col_wardid] in lst_topward:
            df_graph_new.at[df_data.loc[idx,col_numweek],df_data.loc[idx,col_wardid]]=df_data.loc[idx,col_numcase]
        else:
            pass
    return df_graph_new

def assign_numcaseforOthers(df_graph=pd.DataFrame(), df_others=pd.DataFrame(), lst_topward=[], col_index="", col_numweek="", col_numcase=""):
    df_graph["Other wards"] = 0
    for idx in df_others.index:
        df_graph.at[df_others.loc[idx,col_numweek],"Other wards"]=df_others.loc[idx,col_numcase]
    df_graph = df_graph.set_index(col_index)
    return df_graph.loc[:,lst_topward+["Other wards"]]

#Selecting dataframe by profile and ddding values of the passed ward/profile to which df for plotting
def select_dfbylist(df=pd.DataFrame(), lst_lookingfor=[], col_lookingfor="", lst_groupby=[]):
    df_select = pd.DataFrame()
    if len(lst_lookingfor) > 0:
        df_select = df.loc[df[col_lookingfor].isin(lst_lookingfor)].groupby(by=lst_groupby).sum().reset_index()
    else:
        df_select = df
    return df_select
