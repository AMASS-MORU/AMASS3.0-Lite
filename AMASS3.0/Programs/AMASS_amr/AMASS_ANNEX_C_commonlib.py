
#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** CONST FILE and Configurations                                                                   ***#
#***-------------------------------------------------------------------------------------------------***#
# @author: CHALIDA RAMGSIWUTISAK
# Created on: 01 SEP 2023 
import pandas as pd
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
import AMASS_ANNEX_C_const as ACC
import AMASS_amr_commonlib as AL
import AMASS_amr_commonlib_report as REP_AL
import datetime #for setting date-time format
import logging #for creating logfile

#######################################################################
#######################################################################
#####Step1. Merged Hospital-Microbiology data to SaTScan's output######
#######################################################################
#######################################################################
def prepare_fromHospMicro_toSaTScan(logger,df=pd.DataFrame()):
    AL.printlog("Start Cluster signal identification (ANNEX C): " + str(datetime.datetime.now()),False,logger)
    #All specimens
    df_dedup_all = prepare_hospMicroData_profile_byorg(logger,df=df, sh_spc="all", lst_col_ris=[AC.CONST_NEWVARNAME_PREFIX_RIS+col for col in AC.list_antibiotic], filename_profile=ACC.CONST_CURRENTPATH+ACC.CONST_FILENAME_PROFILE)
    prepare_hospMicroData_export(logger,df=df_dedup_all[0], d_profile=df_dedup_all[1], filename_dedup=ACC.CONST_CURRENTPATH+ACC.CONST_FILENAME_HO_DEDUP, sh_spc="all")
    #Blood
    df_dedup_blo = prepare_hospMicroData_profile_byorg(logger,df=df.loc[df[AC.CONST_NEWVARNAME_BLOOD]=="blood"], sh_spc="blo", lst_col_ris=[AC.CONST_NEWVARNAME_PREFIX_RIS+col for col in AC.list_antibiotic], filename_profile=ACC.CONST_CURRENTPATH+ACC.CONST_FILENAME_PROFILE)
    prepare_hospMicroData_export(logger,df=df_dedup_blo[0], d_profile=df_dedup_blo[1], filename_dedup=ACC.CONST_CURRENTPATH+ACC.CONST_FILENAME_HO_DEDUP, sh_spc="blo")

    evaluation_study = retrieve_startEndDate(filename=ACC.CONST_CURRENTPATH+ACC.CONST_FILENAME_REPORT1) #for satscan_param.prm
    for sh_org in ACC.dict_org.keys():
        for sh_spc in ACC.dict_spc.keys():
            print (sh_org, sh_spc)
            prepare_satscanInput(logger,
                                 filename_isolate =ACC.CONST_FILENAME_HO_DEDUP, 
                                 filename_ward    =ACC.CONST_FILENAME_WARD, 
                                 filename_case    =ACC.CONST_FILENAME_INPUT, 
                                 filename_oriparam=ACC.CONST_FILENAME_ORIPARAM, 
                                 filename_newparam=ACC.CONST_FILENAME_NEWPARAM, 
                                 path_satscaninput=ACC.CONST_CURRENTPATH,
                                 sh_org=sh_org, 
                                 sh_spc=sh_spc, 
                                 start_date=evaluation_study[0], 
                                 end_date  =evaluation_study[1])

#####Preparing a merged hospital-microbiology dataframe from the core AMASS#####
################################################################################
#Step1.1.1 Main function for parsing df from core AMASS to the step before grouping isolates by specimen date.
#Hospital-origin, resistant profile, deduplication, AST profiles
#input : filename of HospMicroData
#output 1: dataframe of deduplicated hospital-origin isolates with profiles
#output 2: dictionary for profiles
#output 3: list of passed 
def prepare_hospMicroData_profile_byorg(logger,filename="",df=pd.DataFrame(),lst_col_ris=[],filename_profile="",sh_spc="",col_org=ACC.CONST_NEWVARNAME_ORG3,col_origin =ACC.CONST_COL_ORIGIN, 
                                  col_olddate=ACC.CONST_COL_ORISPCDATE, col_newdate=ACC.CONST_COL_CLEANSPCDATE, 
                                  col_hn     =ACC.CONST_COL_HN,         col_numamr =ACC.CONST_COL_NUMAMR, 
                                  col_profile=ACC.CONST_COL_PROFILE,    dict_ast   =ACC.dict_ast,  dformat=""):
    # df = pd.read_csv(filename).fillna("").replace("aND","-")
    # df = df.fillna("NA").replace("NA","-")
    df_dedup = pd.DataFrame()
    dict_profile_all = {}
    lst_col_ris_passresist = []
    for keys in dict_ast.keys():
        #selecting organisms
        df_str = df.copy()
        df_str[[col_org]+lst_col_ris] = df_str[[col_org]+lst_col_ris].astype(str).fillna("").replace("","-")
        df_str[[col_origin]] = df_str[[col_origin]].astype(int)
        df_org    = select_dfbyOrganism(logger,df=df_str, col_org=col_org, str_selectorg=keys)
        #selecting hospital-origin
        df_org_ho = select_ho(logger,df=df_org, col_origin=col_origin, num_ho=1)
        for val in dict_ast[keys]:
            df_org_ho_r_date           = pd.DataFrame()
            df_org_ho_r_date_sort      = pd.DataFrame()
            df_org_ho_r_date_sort_dedup= pd.DataFrame()
            df_org_ho[val[0]] = df_org_ho[val[0]].astype(str)
            df_org_ho_r = select_resistantProfile(df=df_org_ho, d_ast_val=val)
            if len(df_org_ho_r)>0:
                #datetime formatting for deduplication
                try:
                    df_org_ho_r_date = AL.fn_clean_date(df_org_ho_r,col_olddate,col_newdate,dformat,logger)
                except Exception as e:
                    AL.printlog("Error, ANNEX C formattind Date-Time: " +  str(e),True,logger)

                #sorting datetime for deduplication
                try:
                    df_org_ho_r_date_sort = df_org_ho_r_date.sort_values(by=[col_hn,col_newdate,col_numamr], ascending=[True,True,False], na_position="last")
                    #deduplication
                    df_org_ho_r_date_sort_dedup = df_org_ho_r_date_sort.drop_duplicates(col_hn, keep="first")
                except Exception as e:
                    AL.printlog("Error, ANNEX C Deduplication: " +  str(e),True,logger)
                #selecting profiles from resistant rate
                lst_col_ris_passresist = select_profilebyResistantRate(df=df_org_ho_r_date_sort_dedup, lst_col_ris=lst_col_ris)
                #profile preparation
                df_org_ho_r_date_sort_dedup_profile = create_risprofile(df=df_org_ho_r_date_sort_dedup, lst_col_ris=lst_col_ris_passresist, col_profile=col_profile)
                lst_uniq_profile=drop_dupProfile(df=df_org_ho_r_date_sort_dedup_profile, col_profile=ACC.CONST_COL_PROFILE).tolist()
                d_profile=create_dictforMapProfileID_byorg(lst_profile=lst_uniq_profile, prefixID=ACC.CONST_PRENAME_PROFILEID+val[-1]+"_"+sh_spc+"_")
                dict_profile_all.update(d_profile)
                #concatination
                df_dedup = pd.concat([df_dedup,df_org_ho_r_date_sort_dedup_profile],axis=0)
                print (keys,ACC.CONST_PRENAME_PROFILEID+val[-1]+"_",val,val[-1],len(df_org_ho_r_date_sort_dedup_profile),len(df_dedup))
                #profile exportation
                export_dictforProfile(d_profile=d_profile, lst_col_ris=lst_col_ris_passresist, col_prefixID=ACC.CONST_COL_PROFILEID, col_profile=ACC.CONST_COL_PROFILE, filename_profile=filename_profile+"_"+ACC.dict_org[val[-1]][1]+"_"+ACC.dict_spc[sh_spc])
    df_dedup_profile = df_dedup.copy().reset_index().drop(columns=["index"])
    return df_dedup_profile,dict_profile_all,lst_col_ris_passresist
# def prepare_hospMicroData_profile(logger,filename="",lst_col_ris=[],col_org=ACC.CONST_NEWVARNAME_ORG3,col_origin =ACC.CONST_COL_ORIGIN, 
#                                   col_olddate=ACC.CONST_COL_ORISPCDATE, col_newdate=ACC.CONST_COL_CLEANSPCDATE, 
#                                   col_hn     =ACC.CONST_COL_HN,         col_numamr =ACC.CONST_COL_NUMAMR, 
#                                   col_profile=ACC.CONST_COL_PROFILE,    dict_ast   =ACC.dict_ast):
#     df = pd.read_csv(filename).fillna("").replace("aND","-")
#     df_dedup = pd.DataFrame()
#     for keys in dict_ast.keys():
#         #selecting organisms
#         df_org    = select_dfbyOrganism(df=df, col_org=col_org, str_selectorg=keys)
#         #selecting hospital-origin
#         df_org_ho = select_ho(logger,df=df_org, col_origin=col_origin, num_ho=1)
#         for val in dict_ast[keys]:
#             df_org_ho_r_date           = pd.DataFrame()
#             df_org_ho_r_date_sort      = pd.DataFrame()
#             df_org_ho_r_date_sort_dedup= pd.DataFrame()
#             df_org_ho_r = select_resistantProfile(df=df_org_ho, d_ast_val=val)
#             if len(df_org_ho_r)>0:
#                 #datetime formatting for deduplication
#                 df_org_ho_r_date = AL.fn_clean_date(df=df_org_ho_r, oldfield=col_olddate, cleanfield=col_newdate)
#                 #sorting datetime for deduplication
#                 df_org_ho_r_date_sort = df_org_ho_r_date.sort_values(by=[col_hn,col_newdate,col_numamr], ascending=[True,True,False], na_position="last")
#                 #deduplication
#                 df_org_ho_r_date_sort_dedup = df_org_ho_r_date_sort.drop_duplicates(col_hn, keep="first")
#                 df_dedup = pd.concat([df_dedup,df_org_ho_r_date_sort_dedup],axis=0)
#                 print (keys,val,len(df_org_ho_r_date_sort_dedup),len(df_dedup))
#     df_dedup_profile = df_dedup.copy().reset_index().drop(columns=["index"])
#     df_dedup_profile = create_risprofile(df=df_dedup_profile, lst_col_ris=lst_col_ris, col_profile=col_profile)
#     return df_dedup_profile
#Step1.1.2 Main function for parsing df from core AMASS to the step before grouping isolates by specimen date.
#Mapping profile IDs, dataframe exportation
#input : dataframe of deduplicated hospital-origin isolates with profiles
#output: excel file of dataframe of deduplicated hospital-origin isolates with profiles
def prepare_hospMicroData_export(logger,df=pd.DataFrame(), d_profile={}, filename_dedup="", sh_spc="", 
                                 dict_ast=ACC.dict_ast, col_profile=ACC.CONST_COL_PROFILE, col_profileid=ACC.CONST_COL_PROFILEID, col_org=ACC.CONST_NEWVARNAME_ORG3):
    #mapping profile
    df[col_profileid] = df[col_profile].map(d_profile).fillna("")
    for keys in dict_ast.keys():
        #selecting organisms
        df_org = select_dfbyOrganism(logger,df=df, col_org=col_org, str_selectorg=keys)
        for val in dict_ast[keys]:
            df_org_r = select_resistantProfile(df=df_org, d_ast_val=val)
            if len(df_org_r)>0:
                df_org_r.to_excel(filename_dedup+"_"+ACC.dict_org[val[-1]][1]+"_"+ACC.dict_spc[sh_spc]+".xlsx")
                print (keys,val,len(df_org_r))
#########################Prepare_SaTScan_input.csv##############################
################################################################################
#Step1.2 Main function for parsing df from the step before grouping isolates by specimen date to satscan_input.csv, satscan_location.csv, and satscan_param.prm.
#Grouping isolates by specimen collection date, prepare location and prm files
#input : excel file of dataframe of deduplicated hospital-origin isolates with profiles
#output: csv files for satscan_input.csv, satscan_location.csv, and satscan_param.prm
def prepare_satscanInput(logger,filename_isolate="", filename_ward="", filename_case="", filename_oriparam="", filename_newparam="", path_satscaninput="",
                         sh_org="", sh_spc="", start_date="", end_date="", str_ward=ACC.CONST_VALUE_WARD, 
                         col_amass=ACC.CONST_COL_AMASSNAME,    col_user=ACC.CONST_COL_USERNAME,    col_resist=ACC.CONST_COL_RESISTPROFILE,
                         col_mrsa=ACC.CONST_NEWVARNAME_ASTMRSA,col_3gc=ACC.CONST_NEWVARNAME_AST3GC,col_crab=ACC.CONST_NEWVARNAME_ASTCBPN,  col_van=ACC.CONST_NEWVARNAME_ASTVAN,
                         col_org=ACC.CONST_NEWVARNAME_ORG3,    col_oriorg=ACC.CONST_COL_ORGNAME,   col_profile=ACC.CONST_COL_PROFILEID, 
                         col_wardid=ACC.CONST_COL_WARDID,      col_wardname=ACC.CONST_COL_WARDNAME,col_mappedward=ACC.CONST_COL_MAPPEDWARD,col_cleandate=ACC.CONST_COL_CLEANSPCDATE, 
                         col_clustercode=ACC.CONST_COL_CLUSTERCODE, col_numcase=ACC.CONST_COL_CASE):
    try:
        df = pd.read_excel(path_satscaninput+filename_isolate+"_"+ACC.dict_org[sh_org][1]+"_"+ACC.dict_spc[sh_spc]+".xlsx").fillna("")
        ##ward with group name
        df_ward = read_dictionary(prefix_filename=path_satscaninput+filename_ward)
        df_ward.columns = [col_amass,col_user]
        df_ward = df_ward.loc[df_ward[col_user]!=""] #selecting only defined group
        b_ward = False
        try:
            b_ward = check_ward(df_col=list(df.columns), str_amass_ward=retrieve_uservalue(df=df_ward, amass_name=str_ward, col_amass=col_amass, col_user=col_user)) #boolean wardname
        except Exception as e:
            AL.printlog("Error, ANNEX C Loading dictionary_for_ward: " +  str(e),True,logger)
        d_ward = create_dictforward(df_ward=df_ward, df_micro=df, b_ward=b_ward, col_ward_amass=col_amass, col_ward_user=col_user, col_micro_ward=retrieve_uservalue(df=df_ward, amass_name=str_ward, col_amass=col_amass, col_user=col_user))
        #mapping ward information
        if b_ward:
            const_user_ward = get_keysfromvalues(str_finder=str_ward, d_=d_ward)
            dict_ward_all   = pd.DataFrame()
            try:
                df[col_mappedward] = df[const_user_ward].iloc[:,0].astype(str).map(d_ward)
            except:
                df[col_mappedward] = df[const_user_ward].astype(str).map(d_ward)
            #create dict_ward_all (including ward id and non-ward id)
            dict_ward_all = pd.DataFrame(index=range(len(d_ward.keys())),columns=[col_wardid,col_wardname])
            idx = 0
            for keys in d_ward.keys():
                dict_ward_all.at[idx,col_wardname] = keys
                dict_ward_all.at[idx,col_wardid]   = d_ward[keys]
                idx = idx+1
            df = df.loc[:,[const_user_ward,col_mappedward,col_oriorg,col_org,col_cleandate,col_profile,col_mrsa,col_3gc,col_crab,col_van]]
        else:
            df[col_mappedward] = "ward_000"
            dict_ward_all = pd.DataFrame([["ward_000","no ward information"]],columns=[col_wardid,col_wardname])
            df = df.loc[:,[col_mappedward,col_oriorg,col_org,col_cleandate,col_profile,col_mrsa,col_3gc,col_crab,col_van]]

        d_org = ACC.dict_org[sh_org][0] #get full org name i.e. 'organism_acinetobacter_baumannii'
        df_ns = pd.DataFrame(columns=[col_clustercode,col_cleandate,col_numcase])
        for value in ACC.dict_ast[d_org]:
            if sh_org in value:
                print (value)
                #select organism
                df_org = select_dfbyOrganism(logger,df=df, col_org=col_org, str_selectorg=d_org)
                #select resistant profile
                df_temp = select_resistantProfile(df=df_org, d_ast_val=value)
                df_temp[col_resist] = value[-1] #put short resistant name i.e. crab
                ##Grouping records by cluster_code and specimen_collection_date
                df_temp = create_clusterCode(df=df_temp, check_ward=b_ward, col_clustercode=col_clustercode, col_ward=col_mappedward, col_org=col_resist, col_profile=col_profile)
                df_temp[ACC.CONST_COL_CASE] = 1
                df_temp_code = df_temp.groupby(by=[col_clustercode,col_cleandate]).sum().reset_index()
                df_ns = pd.concat([df_ns,df_temp_code],axis=0)
        try:
            df_ns = df_ns.reset_index().loc[:,[col_clustercode,col_numcase,col_cleandate]]
            df_ns = df_ns.sort_values(by=[col_cleandate]).reset_index().drop(columns=["index"])
        except Exception as e:
            AL.printlog("Error, ANNEX C Case file preparation: " +  str(e),True,logger)
        try:
            df_ns.to_csv(path_satscaninput+filename_case+"_"+sh_org+"_"+sh_spc+".csv",encoding="ascii",index=False)
        except Exception as e:
            AL.printlog("Error, ANNEX C SaTScan input exportation: " +  str(e),True,logger)
    except Exception as e:
        AL.printlog("Error, ANNEX C SaTScan input preparation: " +  str(e),True,logger)
    #For satscan_location.csv
    try:
        prepare_locfile(lst_clustercode=df_ns[col_clustercode].unique(), df_ward_ori=dict_ward_all, col_clustercode=col_clustercode, col_wardid=col_wardid, col_wardname=col_wardname, col_profileid=col_profile, col_resistprofile=col_resist, sh_org=sh_org, sh_spc=sh_spc, path_satscaninput=path_satscaninput)
    except Exception as e:
        AL.printlog("Error, ANNEX C Location file preparation: " +  str(e),True,logger)
    #For satscan_param.prm
    try:
        prepare_prmfile(ori_prmfile=filename_oriparam, new_prmfile=filename_newparam, 
                        start_date=start_date, end_date=end_date, path_satscaninput=path_satscaninput,sh_org=sh_org, sh_spc=sh_spc)
    except Exception as e:
        AL.printlog("Error, ANNEX C Parameter file preparation: " +  str(e),True,logger)


#####Small functions in the step1.1.1-1.1.2#####
#selecting organisms
def select_dfbyOrganism(logger, df=pd.DataFrame(), col_org="", str_selectorg=""):
    df_new = pd.DataFrame()
    try:
        df_new = df.loc[(df[col_org]==str_selectorg)]
    except Exception as e:
        AL.printlog("Error, ANNEX C Selecting organism: " +  str(e),True,logger)
    # if str_selectorg == "organism_enterococcus_faecalis":
    #     df_new = df.loc[(df["organism"]=="Enterococcus faecalis") | (df["organism"]=="efa")]
    # elif str_selectorg == "organism_enterococcus_faecium":
    #     df_new = df.loc[(df["organism"]=="Enterococcus faecium")]
    # else:
    #     df_new = df.loc[(df[col_org]==str_selectorg)]
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
        while (i < len(d_ast_val)) and (i+1!=len(d_ast_val)): #in the case; ["AST3gc","1","ASTCarbapenem","0","3gcr-csec"] >>> do not retrieve "3gcr-crec"
            df_ast = df_ast.loc[(df[d_ast_val[i]]==str(d_ast_val[i+1]))]
            i+=2
    return df_ast
#Creating "R---I-S---R..."
def create_risprofile(df=pd.DataFrame(), lst_col_ris=[], col_profile=""):
    df[col_profile] = ""
    for col in lst_col_ris:
        df[col_profile] = df[col_profile]+ df[col]
    return df
#Droping RIS duplicated sequences
def drop_dupProfile(df=pd.DataFrame(), col_profile=""):
    return df[col_profile].drop_duplicates(keep="first")
#Creating dictionary for mappingID
#d_profile["profile_sequence"]="profile_ID"
def create_dictforMapProfileID_byorg(lst_profile=[], prefixID=""):
    d_profile = {}
    for i in range(len(lst_profile)):
        d_profile[lst_profile[i]] = prefixID+str(i+1) #profile should start from 1
    return d_profile
def create_dictforMapProfileID(lst_profile=[], prefixID=""):
    d_profile = {}
    for i in range(len(lst_profile)):
        d_profile[lst_profile[i]] = prefixID+str(i+1) #profile should start from 1
    return d_profile
#Creating profile information.xlsx
#column profile_ID : profile_001, profile_002, profile_003, ..., profile_NNN
#column RISdrug1 : -, -, R, S
#column RISdrug2 : R, -, -, -
#column RISdrugN : -, -, R, S
def export_dictforProfile(d_profile={}, lst_col_ris=[], col_prefixID="", col_profile="", filename_profile=""):
    d_profile_new = {}
    for keys in d_profile.keys():
        d_profile_new[keys] = d_profile[keys]
    df_profile = pd.DataFrame(data=d_profile_new, index=[col_prefixID]).T.reset_index().set_index([col_prefixID]).rename(columns={"index":col_profile})
    df_profile[lst_col_ris] = ""
    for idx in df_profile.index:
        lst_profile = [*df_profile.loc[idx,col_profile]] #split every string from ["R--S-"] to ["R","-","-","S","-"]
        for j in range(len(df_profile.columns)):
            if df_profile.columns[j] != col_profile:
                col = df_profile.columns[j]
                df_profile.at[idx,col] = lst_profile[j-1]
            else:
                pass
    df_profile.to_excel(filename_profile+".xlsx")
def select_profilebyResistantRate(df=pd.DataFrame(), lst_col_ris=[], 
                                  resistant=ACC.dict_ris["resistant"],intermediate=ACC.dict_ris["intermediate"],susceptible=ACC.dict_ris["susceptible"]):
    lst_col_ris_select = []
    for col in lst_col_ris:
        num_r = int(df[col].str.count(resistant).sum())
        num_i = int(df[col].str.count(intermediate).sum())
        num_s = int(df[col].str.count(susceptible).sum())
        max_r = 0
        min_r = 0
        try:
            max_r = float(ACC.dict_configuration_profile["max_resistant_rate"])
        except:
            max_r = 99.9
        try:
            min_r = float(ACC.dict_configuration_profile["min_resistant_rate"])
        except:
            min_r = 0.1
        denominator = num_r+num_i+num_s
        if denominator>0:
            if   (max_r>0)&(num_r*100/denominator>=max_r):
                pass
            elif (min_r>0)&(num_r*100/denominator<=min_r):
                pass
            else:
                print ("passed",col, num_r,num_i,num_s,denominator,len(df)-denominator)
                lst_col_ris_select.append(col)
    return lst_col_ris_select


#####Small functions in the step1.2#####
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
#Checking ward col name
def check_ward(str_amass_ward="", df_col=[]):
    check_ward = False
    if str_amass_ward == "":
        pass
    else:
        if str_amass_ward in df_col:
            
            check_ward = True
        else:
            pass
    return check_ward
#Creatingdictionary for ward {user_ward_name:amass_ward_name}
def create_dictforward(df_ward=pd.DataFrame(), df_micro=pd.DataFrame(), b_ward=False,
                       col_micro_ward="",      col_ward_amass="",       col_ward_user=""):
    d_ward = {}
    if b_ward:
        ##ward with group name
        for idx in df_ward.index:
            try:
                d_ward[str(int(df_ward.loc[idx,col_ward_user]))] = str(df_ward.loc[idx,col_ward_amass]) #for numeric
            except:
                d_ward[str(df_ward.loc[idx,col_ward_user])] = str(df_ward.loc[idx,col_ward_amass])
        ##ward without group name
        lst_ward_nogroup = []
        try:
            lst_ward_nogroup = list(df_micro[col_micro_ward].iloc[:,0].unique())-d_ward.keys() #if df has duplicated columns >>> select first column
        except:
            lst_ward_nogroup = list(df_micro[col_micro_ward].unique())-d_ward.keys()
        x = 999
        for i in lst_ward_nogroup:
            d_ward[str(i)] = ACC.CONST_VALUE_WARD+"_"+str(x)
            x = x-1
    else:
        pass
    return d_ward
#Finding a dictionary key from value
def get_keysfromvalues(str_finder="", d_={}):
    for keys,values in d_.items():
        if values == str_finder:
            return keys
#create a cluster code column
def create_clusterCode(df=pd.DataFrame(), check_ward=False, col_clustercode="", col_ward="", col_org="", col_profile=""):
    if check_ward:
        df[col_clustercode] = df[col_ward] + ";" + df[col_profile] + ";" + df[col_org]
    else:
        df[col_clustercode] = df[col_ward] + ";" + df[col_profile] + ";" + df[col_org]
    return df
#Prepare parameter file
def prepare_prmfile(ori_prmfile="",start_date="", end_date="", new_prmfile="", sh_org="", sh_spc="", path_satscaninput=""):
    str_wholefile = ""
    file = open(ori_prmfile,"r", encoding="ascii")
    f = file.readline()
    while f != "":  
        for keys,values in ACC.dict_configuration_prm.items():
            if keys in f:
                if ("CaseFile=" == keys) or ("CoordinatesFile=" == keys) or ("ResultsFile=" == keys):
                    f = keys+path_satscaninput+values+sh_org+"_"+sh_spc+".csv" + "\n"
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
    filename_prm = open(path_satscaninput+new_prmfile+"_"+sh_org+"_"+sh_spc+".prm", "w")
    filename_prm.write("".join(str_wholefile))
    filename_prm.close()
#Prepare location file
def prepare_locfile(lst_clustercode=[], df_ward_ori=pd.DataFrame(),
                    col_clustercode="", col_wardid="", col_wardname="", col_profileid="", col_resistprofile="", 
                    sh_org="", sh_spc="", path_satscaninput=""):
    df = pd.DataFrame(lst_clustercode,columns=[col_clustercode]).reset_index().drop(columns=["index"])
    df[["x_axis","y_axis"]]=0
    for idx in df.index:
        df.at[idx,"x_axis"]=idx*3
        df.at[idx,"y_axis"]=idx*3
    df[col_clustercode+"_imd"]=df[col_clustercode].str.split(";")
    df[col_wardid]            =df[col_clustercode+"_imd"].str[0]
    df[col_profileid]         =df[col_clustercode+"_imd"].str[1]
    df[col_resistprofile]     =df[col_clustercode+"_imd"].str[2]
    df   = df.drop(columns=[col_clustercode+"_imd"])
    df_2 = df_ward_ori.merge(df, left_on=col_wardid, right_on=col_wardid,how="inner")
    df_2.loc[:,[col_clustercode,col_wardname]].to_csv(path_satscaninput+ACC.CONST_FILENAME_WARDINFO+"_"+ACC.dict_org[sh_org][1]+"_"+ACC.dict_spc[sh_spc]+".csv",encoding="ascii",index=False)
    df_2.loc[:,[col_clustercode,"x_axis","y_axis"]].to_csv(path_satscaninput+ACC.CONST_FILENAME_LOCATION+"_"+sh_org+"_"+sh_spc+".csv",encoding="ascii",index=False)


#######################################################################
#######################################################################
#############Step2. SaTScan's output to visualized data ###############
#######################################################################
#######################################################################




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

def retrieve_databycondition(df=pd.DataFrame(),col_pval="",pval=0.05):
    return df.loc[df[col_pval].astype(float)<pval,:]

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

def plot(df_org=pd.DataFrame(), palette=[], filename="",
         xlabel="Number of week (Start of weekday)",ylabel="Number of patients(n)*",
         min_yaxis=0,max_yaxis=5,step=2):
    import matplotlib.pyplot as plt
    import numpy as np
    plt.figure()
    df_org.plot(kind='bar', 
            stacked =True, 
            figsize =(20, 10), 
            color =palette, 
            fontsize=16)
    plt.legend(prop={'size': 14},ncol=4)
    plt.ylabel(ylabel, fontsize=16)
    plt.xlabel(xlabel, fontsize=16)
    plt.yticks(np.arange(min_yaxis,max_yaxis+1,step=step), fontsize=16)
    plt.savefig(filename, format='png',dpi=180,transparent=True, bbox_inches="tight")

def get_listofWardwithHighPatient(df=pd.DataFrame(), lst_col_tofind=[], lst_col_togroup=[], lst_col_tosort=[], col_toget="", ascending=False):
    return df.loc[:,lst_col_tofind].groupby(lst_col_togroup).sum().sort_values(lst_col_tosort,ascending=ascending).reset_index().loc[:,col_toget].tolist()

def report_context(c,context_list,pos_x,pos_y,wide,height,font_size=10,font_align=TA_JUSTIFY,line_space=18,left_indent=0):
    context_list_style = []
    style = ParagraphStyle('normal',fontName='Helvetica',leading=line_space,fontSize=font_size,leftIndent=left_indent,alignment=font_align)
    for cont in context_list:
        cont_1 = Paragraph(cont, style)
        context_list_style.append(cont_1)
    f = Frame(pos_x,pos_y,wide,height,showBoundary=0)
    return f.addFromList(context_list_style,c)



#####Supplementary report (Table)#####
def prepare_table(df=pd.DataFrame()):
    return [df.columns.tolist()] + df.values.tolist()
def style_profile(lst_df=[], lst_colwidth=None):
    return Table(lst_df,style=[('FONT',(0,0),(-1,0),'Helvetica-Bold'),
                               ('FONT',(0,1),(-1,-1),'Helvetica'),
                               ('FONTSIZE',(0,0),(-1,0),8),
                               ('FONTSIZE',(0,1),(-1,-1),8),
                               ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                               ('ALIGN',(0,0),(-1,-1),'CENTER'),
                               ('VALIGN',(0,0),(-1,-1),'MIDDLE')], colWidths=lst_colwidth)
#get list of column width for Table style corresponding with no. of column
def style_colwidth(lst_df=[], lst_colwidth=[]):
    return lst_colwidth[:len(lst_df[0])]
#create a list contains column names of each slice dataframe
#[['profile ID', 'DOR', 'CFP', 'TIC', 'PTZ', 'IMP'],['profile ID', 'MEM', 'LEV', 'CFT', 'CIP', 'AMK'],[...]]
def prepare_slicecol(df=pd.DataFrame(), num_col=5):
    lst_col = []
    lst_col_final = []
    b_col=False
    if len(df.columns) >= int(num_col): #get boolean
        b_col = True
    else:
        pass

    if b_col: #use boolean
        count = 0
        col_drugs   = df.columns[1:]
        while count < len(df.columns):
            if count%int(num_col)==0:
                if count+int(num_col) <= len(col_drugs): #page1, page2, page3, page4, ...
                    lst_col = [df.columns[0]]+list(col_drugs[count:count+int(num_col)])
                    lst_col_final.append(lst_col)
                else:#last page
                    lst_col = [df.columns[0]]+list(col_drugs[count:len(df.columns)])
                    if len(lst_col) > 1: #['profile ID', 'MEM', ...] >>> collect the list
                        lst_col_final.append(lst_col)
                    else: #['profile ID'] >>> drop the list
                        pass
            else:
                pass
            count += 1
    else:
        col_drugs   = df.columns[1:]
        lst_col = [df.columns[0]]+list(col_drugs[0:len(df.columns)])
        lst_col_final.append(lst_col)
    return lst_col_final
#create a list contains pairs of start-end of index of each slice dataframe
#[[0,4],[5,9],[10,14],...]
def prepare_slicerow(df=pd.DataFrame(), num_row=5):
    b_row=False
    if len(df) >= int(num_row): #get boolean
        b_row = True
    else:
        pass
    lst_row_final = []
    if b_row: #use boolean
        count = 0
        while count < len(df):
            if count%int(num_row)==0:
                if count+int(num_row) <= len(df): #page1,page2, page3, page4, ...
                    lst_row_final.append([count,count+int(num_row)-1])
                else: #last page
                    lst_row_final.append([count,len(df)])
            else:
                pass
            count += 1
    else:
        lst_row_final.append([0,len(df)-1])
    return lst_row_final

def report_context(c,context_list,pos_x,pos_y,wide,height,font_size=10,font_align=TA_JUSTIFY,line_space=18,left_indent=0):
    context_list_style = []
    style = ParagraphStyle('normal',fontName='Helvetica',leading=line_space,fontSize=font_size,leftIndent=left_indent,alignment=font_align)
    for cont in context_list:
        cont_1 = Paragraph(cont, style)
        context_list_style.append(cont_1)
    f = Frame(pos_x,pos_y,wide,height,showBoundary=0)
    return f.addFromList(context_list_style,c)

def get_yaxisforportrait(df=pd.DataFrame(),yaxis=0.5):
    new_yaxis = 0.0
    if (len(df) < 5):
        new_yaxis = 5.5
    elif (len(df) >= 5) and (len(df) < 10):
        new_yaxis = 5.0
    elif (len(df) >= 10) and (len(df) < 15):
        new_yaxis = 4.5
    elif (len(df) >= 16) and (len(df) < 20):
        new_yaxis = 3.0
    elif (len(df) >= 21) and (len(df) < 25):
        new_yaxis = 2.3
    else:
        new_yaxis = yaxis
    return new_yaxis

def get_yaxisforlandscape(df=pd.DataFrame(),yaxis=0.5):
    new_yaxis = 0.0
    if (len(df) < 5):
        new_yaxis = 3.0
    elif (len(df) >= 5) and (len(df) < 10):
        new_yaxis = 2.0
    elif (len(df) >= 10) and (len(df) < 15):
        new_yaxis = 1.8
    elif (len(df) >= 16) and (len(df) < 20):
        new_yaxis = 1.5
    elif (len(df) == 20):
        new_yaxis = 0.5
    else:
        new_yaxis = yaxis
    return new_yaxis

def create_pdf_table(c, b_portraits=True, df=pd.DataFrame(), num_row=35, lst_colwidth=None,  title_color="#3e4444",
                     title_name="",   title_fontsize=16, title_xaxis=1.07, title_yaxis=10.5, 
                     table_width=500, table_height=300,  table_xaxis=0.5,  table_yaxis=1.0):
    lst_df = prepare_table(df = df)
    REP_AL.report_title(c,title_name,float(title_xaxis)*inch,float(title_yaxis)*inch,font_color=title_color,font_size=int(title_fontsize))
    if lst_colwidth is None:
        table_draw = style_profile(lst_df, lst_colwidth=None)
    else:
        table_draw = style_profile(lst_df, lst_colwidth=style_colwidth(lst_df=lst_df, lst_colwidth=lst_colwidth))
    table_draw.wrapOn(c, int(table_width), int(table_height))
    if b_portraits:
        table_draw.drawOn(c, float(table_xaxis)*inch, float(get_yaxisforportrait(df=df,yaxis=table_yaxis))*inch)
    else:
        table_draw.drawOn(c, float(table_xaxis)*inch, float(get_yaxisforlandscape(df=df,yaxis=table_yaxis))*inch)
    c.showPage()