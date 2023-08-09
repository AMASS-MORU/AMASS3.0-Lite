#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** MAIN ANALYSIS (SECTION 1-6, ANEX A                                                              ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate AMR surveillance reports, Supplementary data indicators reports, and Data verification logfile reports systematically.
# Code is rewrite from R code developed by Cherry Lim in AMASS version 2.0.15
# @author: PRAPASS WANNAPINIJ
# Created on: 20 Oct 2022

import os
#import csv
#import logging #for creating error_log
import pandas as pd #for creating and manipulating dataframe
import numpy as np #for creating arrays and calculations
from datetime import datetime
#from AMASS_3_amr_commonlib import * #for importing amr functions
#from xlsx2csv import Xlsx2csv

import calendar as cld
import psutil,gc
import AMASS_amr_const as AC
import AMASS_amr_commonlib as AL
import AMASS_amr_analysis_annex_b as ANNEX_B
import AMASS_amr_report_new as AMR_REPORT_NEW
import AMASS_supplementary_report as SUP_REPORT
from scipy.stats import norm # -> must moveto common lib

bisdebug = True


# Function for get lower uper CI -> must moveto common lib
# to calculate lower 95% CI
def fn_wilson_lowerCI(x, n, conflevel, decimalplace):
  zalpha = abs(norm.ppf((1-conflevel)/2))
  phat = x/n
  bound = (zalpha*((phat*(1-phat)+(zalpha**2)/(4*n))/n)**(1/2))/(1+(zalpha**2)/n)
  midpnt = (phat+(zalpha**2)/(2*n))/(1+(zalpha**2)/n)
  lowlim = round((midpnt - bound)*100, decimalplace)
  return lowlim

# to calculate upper 95% CI
def fn_wilson_upperCI(x, n, conflevel, decimalplace):
  zalpha = abs(norm.ppf((1-conflevel)/2))
  phat = x/n
  bound = (zalpha*((phat*(1-phat)+(zalpha**2)/(4*n))/n)**(1/2))/(1+(zalpha**2)/n)
  midpnt = (phat+(zalpha**2)/(2*n))/(1+(zalpha**2)/n)
  uplim = round((midpnt + bound)*100, decimalplace)
  return uplim

# function that print obj if in debug mode (bisdebug = true)
def printdebug(obj) :
    try:
        if bisdebug :
            print(obj)
    except Exception as e: 
        print(e)
def check_config(df_config, str_process_name):
    #Checking process is either able for running or not
    #df_config: Dataframe of config file
    #str_process_name: process name in string fmt
    #return value: Boolean; True when parameter is set "yes", False when parameter is set "no"
    config_lst = df_config.iloc[:,0].tolist()
    result = ""
    if df_config.loc[config_lst.index(str_process_name),"Setting parameters"] == "yes":
        result = True
    else:
        result = False
    return result


def clean_hn(df,oldfield,newfield) :
    df[newfield] = df[oldfield].fillna("").astype("string").str.strip()
    try:
        df[newfield] = df[newfield].str.replace(" ","").str.lower()
        #Case read and add .0 at the end of HN as number 
        df[newfield] = df[newfield].str.replace(r'\.0$','',regex=True)
    except Exception as e: # work on python 3.x
        print(e)
    return df
def caldays(df,coldate,dbaselinedate) :
    return (df[coldate] - dbaselinedate).dt.days
def fn_allmonthname():
    month = 1
    return [cld.month_name[(month % 12 + i) or month] for i in range(12)]
def get_randomstring_data(temp_df,sfield,iRow):
    try:
        sstr = ""
        if len(temp_df) > 0:
            temp_df = temp_df.loc[np.random.permutation(temp_df.index)[:iRow if len(temp_df) > iRow else len(temp_df)]]
            for index, row in temp_df.iterrows():
                sstr = sstr + (', ' if len(sstr) > 0 else '') + str(row[sfield])
        return sstr
    except:
        return ""
def check_hn_format(orgdf,hnfield,logger):
    df2 = pd.DataFrame(columns =["str_length","counts","percent"]) 
    try:
        df = orgdf[[hnfield]].fillna('')
        df["str_length"] = df[hnfield].str.len()
        df2 = df.groupby(["str_length"]).size().reset_index(name='counts')
        """
        iall = df2['counts'].sum()
        df2.sort_values(by=['counts'])
        for i in range(len(df2)) :
        """
        df2["percent"] = round(df2['counts']/df2['counts'].sum(),4)*100
    except Exception as e:
        AL.printlog("Warning : create data varidation log for HN charecter length distribution", False, logger)
        logger.exception(e)
    return df2
def check_date_format(orgdf,oldfield,logger):
    try:
        isalreadydatecol = pd.api.types.is_datetime64_any_dtype(orgdf[oldfield])
    except:
        isalreadydatecol = False
    df = orgdf[[oldfield]]
    df_logformat = pd.DataFrame(columns =["dateformat","exampledate"]) 
    if isalreadydatecol:
        onew_row = {"dateformat":'Dates are in excel standard date format : ',"exampledate":get_randomstring_data(df,oldfield,3)}   
        df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
    else:
        sfield = oldfield + "amasschk"
        cft_1 = sfield + "_amassd1"
        cft_2 = sfield + "_amassd2"
        try:
            df[sfield] = df[oldfield].astype("string")
            df[sfield] = df[sfield].str.replace('-', '/', regex=False)
            df[cft_1] = df[sfield].str.split("/", n = 2, expand = True)[0]
            df[cft_1] = pd.to_numeric(df[cft_1],downcast='signed',errors='coerce').fillna(0)
            df[cft_2] = df[sfield].str.split("/", n = 2, expand = True)[1]
            df[cft_2] = pd.to_numeric(df[cft_2],downcast='signed',errors='coerce').fillna(0)
            temp_df = df.loc[(df[cft_1]>12) & (df[cft_1]<32)]
            if len(temp_df) > 0:
                onew_row = {"dateformat":'DMY',"exampledate":get_randomstring_data(temp_df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
            temp_df = df.loc[(df[cft_2]>12) & (df[cft_2]<32)]
            if len(temp_df) > 0:
                onew_row = {"dateformat":'MDY',"exampledate":get_randomstring_data(temp_df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
            temp_df = df.loc[(df[cft_1]>31)]
            if len(temp_df) > 0:
                onew_row = {"dateformat":'YMD',"exampledate":get_randomstring_data(temp_df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
            temp_df = df.loc[(df[cft_1]==0)]
            if len(temp_df) > 0:
                onew_row = {"dateformat":'Others defined',"exampledate":get_randomstring_data(temp_df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
        except:
            try:
                onew_row = {"dateformat":'Undefinded',"exampledate":get_randomstring_data(df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
            except:
                pass
    return df_logformat
def clean_date(df,oldfield,cleanfield,dformat,logger):
    return AL.fn_clean_date(df,oldfield,cleanfield,dformat,logger)
def clean_date_old_2(df,oldfield,cleanfield,dformat,logger):
    CDATEFORMAT_YMD =["%Y/%m/%d","%y/%m/%d"] 
    CDATEFORMAT_DMY =["%d/%m/%Y","%d/%m/%y"]
    CDATEFORMAT_MDY =["%m/%d/%Y","%m/%d/%y"]
    CDATEFORMAT_OTH =["%d/%b/%Y","%d/%b/%y","%d/%B/%Y","%d/%B/%y"]
    cleanfieldtemp = cleanfield + "_tmpamassf"
    cft_1 = cleanfield + "_d1"
    cft_2 = cleanfield + "_d2"
    if oldfield != cleanfield:
        try:
            isalreadydatecol = pd.api.types.is_datetime64_any_dtype(df[oldfield])
        except:
            isalreadydatecol = False
        if isalreadydatecol != True:
            df[cleanfield] = df[oldfield].astype("string")
            df[cleanfield] = df[cleanfield].fillna("1900-01-01")
            df[cleanfield] = df[cleanfield].str.split(" ", n = 1, expand = True)[0]
            iDMY = 0
            iYMD = 0
            iMDY = 0
            iOTH = 0
            try:
                df[cleanfield] = df[cleanfield].str.replace('-', '/', regex=False)
                df[cft_1] = df[cleanfield].str.split("/", n = 2, expand = True)[0]
                df[cft_1] = pd.to_numeric(df[cft_1],downcast='signed',errors='coerce')
                df[cft_2] = df[cleanfield].str.split("/", n = 2, expand = True)[1]
                df[cft_2] = pd.to_numeric(df[cft_2],downcast='signed',errors='coerce')
                iDMY = len(df[(df[cft_1]>12) & (df[cft_1]<32)])                   
                iYMD = len(df[(df[cft_1]>31)])
                iMDY = len(df[(df[cft_2]>12) & (df[cft_2]<32)])
                df = df.drop(columns=[cft_1])  
                df = df.drop(columns=[cft_2]) 
                print('Count date format DMY:' + str(iDMY) + ', MDY:' + str(iMDY) + ', YMD:' + str(iYMD))
            except Exception as e:
                AL.printlog("Warning date format of " + oldfield + " may be not in convert format defined or in other format", False, logger)
                logger.exception(e)
                iOTH = 1
            df_format = pd.DataFrame({'fname':['YMD','DMY','MDY','Others'],'fcount':[iYMD,iDMY,iMDY,iOTH],'cformat':[CDATEFORMAT_YMD,CDATEFORMAT_DMY,CDATEFORMAT_MDY,CDATEFORMAT_OTH]}) 
            df_format = df_format.sort_values(by=['fcount'],ascending=False)
            df[cleanfieldtemp] = df[cleanfield]
            bfirstformat = True
            for index, row in df_format.iterrows():
                print('Convert data format: ' + row['fname'] )
                for sf in row['cformat']:
                    if bfirstformat:
                        df[cleanfield] = pd.to_datetime(df[cleanfield], format=sf, errors="coerce")
                    else:
                        if df[cleanfield].isnull().values.any() == False:
                            break
                        df[cleanfield] = df[cleanfield].fillna(pd.to_datetime(df[cleanfieldtemp], format=sf, errors="coerce"))
                    bfirstformat = False
            if df[cleanfield].isnull().values.any() == True:
                df[cleanfield] = df[cleanfield].fillna(pd.to_datetime(df[oldfield], errors="coerce"))
            #df.loc[df[cleanfield]<datetime(1900, 1, 1),cleanfield] = np.nan
            df = df.drop(columns=[cleanfieldtemp])  
        else:
            df[cleanfield] = df[oldfield]
            print("Note: " + oldfield + " is already date time data type")
    return df
def clean_date_old_1(df,oldfield,cleanfield,dformat,logger):
    CDATEFORMAT_DMY =["%d/%m/%Y","%d/%m/%y"]
    CDATEFORMAT_MDY =["%m/%d/%Y","%m/%d/%y"]
    CDATEFORMAT_OTH =["%d/%b/%Y","%d/%b/%y","%d/%B/%Y","%d/%B/%y"]
    cleanfieldtemp = cleanfield + "_tmpamassf"
    if oldfield != cleanfield:
        try:
            isalreadydatecol = pd.api.types.is_datetime64_any_dtype(df[oldfield])
        except:
            isalreadydatecol = False
        if isalreadydatecol != True:
            df[cleanfield] = df[oldfield].astype("string")
            df[cleanfield] = df[cleanfield].fillna("1900-01-01")
            df[cleanfield] = df[cleanfield].str.split(" ", n = 1, expand = True)[0]
            df[cleanfield] = df[cleanfield].str.replace('-', '/', regex=False)
            bisDMY= True
            try:
                #check if second value in date split is day -> m-d-y
                if df[cleanfield].str.split("/", n = 2, expand = True)[1].astype(int).max() > 12:
                   bisDMY= False
            except:
                bisDMY= True
            df[cleanfieldtemp] = df[cleanfield]
            df[cleanfield] = pd.to_datetime(df[cleanfield],format="%Y/%m/%d", errors="coerce")
            if bisDMY==True:
                dformat = CDATEFORMAT_DMY + CDATEFORMAT_MDY + CDATEFORMAT_OTH
            else:
                dformat = CDATEFORMAT_MDY + CDATEFORMAT_DMY + CDATEFORMAT_OTH
            for sf in dformat:
                if df[cleanfield].isnull().values.any() == False:
                    break
                df[cleanfield] = df[cleanfield].fillna(pd.to_datetime(df[cleanfieldtemp], format=sf, errors="coerce"))
            if df[cleanfield].isnull().values.any() == True:
                df[cleanfield] = df[cleanfield].fillna(pd.to_datetime(df[oldfield], errors="coerce"))
            #df.loc[df[cleanfield]<datetime(1900, 1, 1),cleanfield] = np.nan
            df = df.drop(columns=[cleanfieldtemp])  
        else:
            df[cleanfield] = df[oldfield]
            print("Note: " + oldfield + " is already date time data type")
    return df
def fn_clean_date_andcalday_year_month(df,oldfield,cleanfield,caldayfield,calyearfield,calmonthfield, CDATEFORMAT, ORIGIN_DATE,logger) :
    try:
        bisok = True
        try:
            if oldfield in df.columns:
                df  = clean_date(df, oldfield, cleanfield, CDATEFORMAT,logger)
                if caldayfield!="" : df[caldayfield] = caldays(df, cleanfield, ORIGIN_DATE)
                if calyearfield!="": df[calyearfield] = df[cleanfield].dt.strftime("%Y")
                if calmonthfield!="": df[calmonthfield] = df[cleanfield].dt.strftime("%B")
            else:
                #raise Exception("No field : " + oldfield)
                AL.printlog("Warning : No field : " + oldfield, False, logger)
                bisok = False
        except Exception as e:
            AL.printlog("Error : unable to convert and calculate day year month for " + oldfield, False, logger)
            logger.exception(e)
            bisok = False
        if bisok == False:
            df[cleanfield] = np.nan
            if caldayfield!="": df[caldayfield] = np.nan
            if calyearfield!="":df[calyearfield] = np.nan
            if calmonthfield!="": df[calmonthfield] = np.nan
        return df  
    except:
        return df      
def fn_notindata(df_source,df_indi,sfield_source,sfield_indi,bindidropdup=False):
    df1=pd.DataFrame()
    df2=pd.DataFrame()
    try:
        df1[sfield_source] =df_source[sfield_source]
        if bindidropdup:
            df2[sfield_indi+"_indi"] = df_indi[sfield_indi].drop_duplicates()
        else:
            df2[sfield_indi+"_indi"] = df_indi[sfield_indi]
        df1 = df1.merge(df2, how="left", left_on=sfield_source, right_on=sfield_indi+"_indi",suffixes=("", "_indi"))
        df1[sfield_indi+"_indi"] = df1[sfield_indi+"_indi"] .fillna("")
        df1=df1[df1[sfield_indi+"_indi"]==""]
    except:
        pass
    return df1
def fn_mergededup_hospmicro(df_micro,df_hosp,bishosp_ava,df_dict,dict_datavaltoamass,dict_inforg_datavaltoamass,dict_gender_datavaltoamass,dict_died_datavaltoamass,logger,df_list_matchrid) :
    df_merged = pd.DataFrame()
    bMergedsuccess = False
    sErrorat = ""
    # Merge if got same column name in both (such as date_of_admission) it will use the date_of-admission from micro dataframe
    if bishosp_ava:
        try:
            #Remove micro hn not in hosp - This part may need to change if need unmatch data
            #df_merged = df_micro[df_micro[AC.CONST_NEWVARNAME_HN].isin(df_hosp[AC.CONST_NEWVARNAME_HN_HOSP])]
            #if len(df_merged) <= 0:
            #    df_merged = df_micro.copy(deep=True)
            #Inner join merge
            sErrorat = "(do merge)"
            df_merged = df_micro.copy(deep=True)
            df_merged = df_merged.merge(df_hosp, how="inner", left_on=AC.CONST_NEWVARNAME_HN, right_on=AC.CONST_NEWVARNAME_HN_HOSP,suffixes=("", AC.CONST_MERGE_DUPCOLSUFFIX_HOSP))
            df_list_matchrid[AC.CONST_NEWVARNAME_MICROREC_ID] =df_merged[AC.CONST_NEWVARNAME_MICROREC_ID].drop_duplicates()
            sErrorat = "(do clean date)"
            df_merged = fn_clean_date_andcalday_year_month(df_merged, AC.CONST_VARNAME_ADMISSIONDATE, AC.CONST_NEWVARNAME_CLEANADMDATE, AC.CONST_NEWVARNAME_DAYTOADMDATE, AC.CONST_NEWVARNAME_ADMYEAR, AC.CONST_NEWVARNAME_ADMMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            df_merged = fn_clean_date_andcalday_year_month(df_merged, AC.CONST_VARNAME_DISCHARGEDATE, AC.CONST_NEWVARNAME_CLEANDISDATE, AC.CONST_NEWVARNAME_DAYTODISDATE, AC.CONST_NEWVARNAME_DISYEAR, AC.CONST_NEWVARNAME_DISMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            df_merged = fn_clean_date_andcalday_year_month(df_merged, AC.CONST_VARNAME_BIRTHDAY, AC.CONST_NEWVARNAME_CLEANBIRTHDATE, AC.CONST_NEWVARNAME_DAYTOBIRTHDATE, "","", AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            #cal end date to get admission period
            df_merged[AC.CONST_NEWVARNAME_DAYTOENDDATE] = df_merged[AC.CONST_NEWVARNAME_DAYTODISDATE]
            df_merged[AC.CONST_NEWVARNAME_DAYTOSTARTDATE] = df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE]
            try:
                calcolmaxdayinclude(df_merged,AC.CONST_NEWVARNAME_CLEANSPECDATE,df_merged,AC.CONST_NEWVARNAME_CLEANADMDATE,AC.CONST_NEWVARNAME_DAYTOENDDATE,AC.CONST_ORIGIN_DATE,"",logger)   
                calcolmindayinclude(df_merged,AC.CONST_NEWVARNAME_CLEANSPECDATE,df_merged,AC.CONST_NEWVARNAME_CLEANADMDATE,AC.CONST_NEWVARNAME_DAYTOSTARTDATE,AC.CONST_ORIGIN_DATE,"",logger)  
            except Exception as e:
                AL.printlog("Warning : calculate min max include date (Merged data)"  + str(e),True,logger)
                logger.exception(e)
                pass
            # Remove unmatched
            sErrorat = "(do remove unmatch spec date vs admission period)"
            #df_merged = df_merged[(df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE]>=df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE]) & (df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE]<=df_merged[AC.CONST_NEWVARNAME_DAYTODISDATE])]
            df_merged = df_merged[(df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE]>=df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE]) & (df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE]<=df_merged[AC.CONST_NEWVARNAME_DAYTOENDDATE])]
            # Translate gender, age year, age cat data
            sErrorat = "(do add columns)"
            # Translate gender, age year, age cat data
            df_merged[AC.CONST_NEWVARNAME_GENDERCAT] = df_merged[AC.CONST_VARNAME_GENDER].astype("string").str.strip().map(dict_gender_datavaltoamass)
            # CO/HO (If want to support case micro data have admission date move it outside of bishosp_ava condition)
            if AC.CONST_VARNAME_COHO in df_merged.columns:
                df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] = df_merged[AC.CONST_VARNAME_COHO].astype("string").str.strip().map(dict_inforg_datavaltoamass)
                df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] = np.select([(df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] == AC.CONST_DICT_COHO_CO),(df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] == AC.CONST_DICT_COHO_HO)],
                                                                            [0,1],
                                                                            default=np.nan)
            else:
                df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] = np.nan
            df_merged[AC.CONST_NEWVARNAME_COHO_FROMCAL] = np.select([((df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE] - df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE]) < 2),
                                                                         ((df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE] - df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE]) >= 2)],
                                                                        [0,1],
                                                                        default=np.nan)     
            if df_dict[df_dict[AC.CONST_DICTCOL_AMASS]==AC.CONST_DICT_COHO_AVAIL][AC.CONST_DICTCOL_DATAVAL].values[0] == "yes" :
               df_merged[AC.CONST_NEWVARNAME_COHO_FINAL] =  df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS]
            else:
               df_merged[AC.CONST_NEWVARNAME_COHO_FINAL] =  df_merged[AC.CONST_NEWVARNAME_COHO_FROMCAL]  
            df_merged[AC.CONST_NEWVARNAME_DISOUTCOME] = df_merged[AC.CONST_VARNAME_DISCHARGESTATUS].astype("string").str.strip().map(dict_died_datavaltoamass).fillna(AC.CONST_ALIVE_VALUE) # From R code line 1154
            sErrorat = "(do change columns to category type)"
            #df_merged[AC.CONST_NEWVARNAME_AGECAT] = df_merged[AC.CONST_NEWVARNAME_AGECAT].astype("category")
            df_merged[AC.CONST_NEWVARNAME_DISOUTCOME] = df_merged[AC.CONST_NEWVARNAME_DISOUTCOME].astype("category")
            bMergedsuccess = True
        except Exception as e: # work on python 3.x
            AL.printlog("Failed merge hosp and micro data on " + sErrorat + " : " + str(e),True,logger)
            logger.exception(e)
            pass
    if (bishosp_ava==False) or (bMergedsuccess==False) :
        try:
            AL.printlog("No hosp data/merged error, merged data is micro data",False,logger)
            df_merged = df_micro.copy(deep=True)
            df_merged[AC.CONST_NEWVARNAME_HN_HOSP] = np.nan
            df_merged[AC.CONST_NEWVARNAME_CLEANADMDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_ADMYEAR] = np.nan
            df_merged[AC.CONST_NEWVARNAME_ADMMONTHNAME] = np.nan
            df_merged[AC.CONST_NEWVARNAME_CLEANDISDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DAYTODISDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DISYEAR] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DISMONTHNAME] = np.nan
            df_merged[AC.CONST_NEWVARNAME_CLEANBIRTHDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DAYTOBIRTHDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_GENDERCAT] = np.nan
            df_merged[AC.CONST_NEWVARNAME_AGEYEAR] = np.nan
            #df_merged[AC.CONST_NEWVARNAME_AGECAT] = np.nan
            df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] = np.nan
            df_merged[AC.CONST_NEWVARNAME_COHO_FROMCAL] = np.nan
            df_merged[AC.CONST_NEWVARNAME_COHO_FINAL] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DISOUTCOME] =np.nan
        except Exception as e: # work on python 3.x
            AL.printlog("Failed case no hosp data/merged error, merged data is micro data : " + str(e),True,logger)
            logger.exception(e)
    return df_merged
def calcolmindayinclude(dfm,scolm,dfh,scolh,scolstartdate,dorgdate,ddefault,logger) :
    try:
        dmin_data_include = dfm[scolm].min()
        dtemp = dfh[scolh].min()
        if dtemp > dmin_data_include:
            dmin_data_include = dtemp
        idaytomin_date_include = (dmin_data_include - dorgdate).days
        AL.printlog("Min date include = " + str(dmin_data_include) + " which is days " + str(idaytomin_date_include),False,logger)
        dfh[scolstartdate] = dfh[scolstartdate].fillna(idaytomin_date_include)
        dfh.loc[dfh[scolstartdate] < idaytomin_date_include, scolstartdate] = idaytomin_date_include
        #reset those adm date is null back to null value
        dfh.loc[dfh[scolh].isnull(), scolstartdate] = np.nan
        return dmin_data_include
    except Exception as e:
        AL.printlog("Warning : Fail to calculate/set value start date data include: " +  str(e),False,logger)
        logger.exception(e)
        return ddefault
def calcolmaxdayinclude(dfm,scolm,dfh,scolh,scolendate,dorgdate,ddefault,logger) :
    try:
        dmax_data_include = dfm[scolm].max()
        dtemp = dfh[scolh].max()
        if dtemp < dmax_data_include:
            dmax_data_include = dtemp
        idaytomax_date_include = (dmax_data_include - dorgdate).days
        AL.printlog("Max date include = " + str(dmax_data_include) + " which is days " + str(idaytomax_date_include),False,logger)
        dfh[scolendate] = dfh[scolendate].fillna(idaytomax_date_include)
        dfh.loc[dfh[scolendate] > idaytomax_date_include, scolendate] = idaytomax_date_include
        #reset those adm date is null back to null value
        dfh.loc[dfh[scolh].isnull(), scolendate] = np.nan
        return dmax_data_include
    except Exception as e:
        AL.printlog("Warning : Fail to calculate/set value end date data include: " +  str(e),False,logger)
        logger.exception(e)
        return ddefault
# Save temp file if in debug mode
def debug_savecsv(df,fname,bdebug,iquotemode,logger)  :
    if bdebug :
        AL.fn_savecsv(df,fname,iquotemode,logger)
        #return True
# Filter orgcat before dedup (For merge data)
def fn_deduplicatebyorgcat_hospmico(df,colname,orgcat) :
    return fn_deduplicatedata(df.loc[df[colname]==orgcat],[AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_AMR,AC.CONST_NEWVARNAME_AMR_TESTED,AC.CONST_NEWVARNAME_CLEANADMDATE],[True,True,False,False,True],"last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
# Filter orgcat before dedup
def fn_deduplicatebyorgcat(df,colname,orgcat) :
    return fn_deduplicatedata(df.loc[df[colname]==orgcat],[AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_AMR,AC.CONST_NEWVARNAME_AMR_TESTED],[True,True,False,False],"last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
# deduplicate data by sort the data by multiple column (First order then second then..) and then select either first or last of each group of the first order and/or second and/or..
def fn_deduplicatedata(df,list_sort,list_order,na_posmode,list_dupcolchk,keepmode) :
    return df.sort_values(by = list_sort, ascending = list_order, na_position = na_posmode).drop_duplicates(subset=list_dupcolchk, keep=keepmode)
def sub_printprocmem(sstate,logger) :
    try:
        process = psutil.Process(os.getpid())
        AL.printlog("Memory usage at state " +sstate + " is " + str(process.memory_info().rss) + " bytes.",False,logger) 
    except:
        AL.printlog("Error get process memory usage at " + sstate,True,logger)
#Ward dictionary function
def fn_clean_ward(df,scol_wardid,scol_wardtype,path,f_dict_ward,logger):
    bOK = False
    try:
        df_dict_ward = AL.readxlsorcsv_noheader(path,f_dict_ward, [AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL,"WARDTYPE","REQ","EXPLAINATION"],logger)
        df_dict_ward = df_dict_ward[df_dict_ward[AC.CONST_DICTCOL_DATAVAL].str.strip() != ""]
        scol_wardorg = df_dict_ward[df_dict_ward [AC.CONST_DICTCOL_AMASS] == AC.CONST_VARNAME_WARD].iloc[0][AC.CONST_DICTCOL_DATAVAL]
        df_dict_ward = df_dict_ward[df_dict_ward[AC.CONST_DICTCOL_AMASS].str.startswith("ward_")]
        if scol_wardorg in df.columns:
            tempdict = pd.Series(df_dict_ward[AC.CONST_DICTCOL_AMASS].values,index=df_dict_ward[AC.CONST_DICTCOL_DATAVAL].str.strip()).to_dict()
            df[scol_wardid] = df[scol_wardorg].astype("string").map(tempdict)
            tempdict = pd.Series(df_dict_ward["WARDTYPE"].values,index=df_dict_ward[AC.CONST_DICTCOL_DATAVAL].str.strip()).to_dict()
            df[scol_wardtype] = df[scol_wardorg].astype("string").map(tempdict)
            df.rename(columns={scol_wardorg:AC.CONST_VARNAME_WARD}, inplace=True)
            bOK = True
        else:
            AL.printlog("Warning : Fail to convert ward data: " + "No ward column as specify in dictionary of ward",False,logger)
    except Exception as e:
        AL.printlog("Warning : Fail to convert ward data: " +  str(e),False,logger)
        logger.exception(e)
    if bOK == False:
        df[scol_wardid] = np.nan
        df[scol_wardtype] = np.nan
    try:
        df[scol_wardid].astype("category")
        df[scol_wardtype].astype("category")
    except:
        pass
    return bOK    
def mainloop() :    
    dict_progvar = {}  
    ## Init log 
    logger = AL.initlogger('AMR anlaysis',"./log_amr_analysis.txt")
    AL.printlog("AMASS version : " + AC.CONST_SOFTWARE_VERSION,False,logger)
    AL.printlog("Pandas library version : " + str(pd.__version__),False,logger)
    AL.printlog("Start AMR analysis: " + str(datetime.now()),False,logger)
    ## Date format
    fmtdate_text = "%d %b %Y"
    try:
        # If folder doesn't exist, then create it.
        if not os.path.isdir(AC.CONST_PATH_RESULT) : os.makedirs(AC.CONST_PATH_RESULT)
        path_repwithPID = AC.CONST_PATH_REPORTWITH_PID
        path_variable = AC.CONST_PATH_VAR
        path_input = "./"
        ## Import data 
        sub_printprocmem("Start main loop",logger)
        df_micro = pd.DataFrame()
        #df_micro_original = pd.DataFrame()
        df_hosp = pd.DataFrame()
        df_hosp_formerge = pd.DataFrame()
        df_config = AL.readxlsxorcsv(path_input +"Configuration/", "Configuration",logger)
        if check_config(df_config, "amr_surveillance_function") :
            #import microbiology
            if AL.checkxlsorcsv(path_input,"microbiology_data_reformatted") :
                df_micro = AL.readxlsxorcsv(path_input,"microbiology_data_reformatted",logger)
            else :
                df_micro = AL.readxlsxorcsv(path_input,"microbiology_data",logger)
            df_micro_annexb = df_micro.copy(deep=True)
        AL.printlog("Succesful read micro data file with " + str(len(df_micro)) + " records",False,logger)  
        #Special task for log ast --------------------------------------------------------------------------------------------------
        #"Save the log_ast.xlsx for data validation file"
        temp_df = pd.DataFrame(columns =["Antibiotics","frequency_raw"]) 
        for scolname in df_micro:
            try:   
                n_row = len(df_micro[(df_micro[scolname].isnull() == False) & (df_micro[scolname] != "")])
                onew_row = {"Antibiotics":scolname,"frequency_raw":str(n_row)}     
                temp_df = pd.concat([temp_df,pd.DataFrame([onew_row])], ignore_index = True)
            except Exception as e: 
                print(e)    
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_ast.xlsx", logger):
            AL.printlog("Warning : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_ast.xlsx",False,logger)  
        del temp_df
        gc.collect()
        #--------------------------------------------------------------------------------------------------------------------------
        #df_micro_original = df_micro.copy(deep=True)
        if AL.checkxlsorcsv(path_input,"hospital_admission_data"):
            df_hosp = AL.readxlsxorcsv(path_input,"hospital_admission_data",logger)
            AL.printlog("Succesful read hospital data file with " + str(len(df_hosp)) + " records",False,logger)  
        else:
            AL.printlog("Note : No hospital data file",False,logger)  
        bishosp_ava = not df_hosp.empty
        #sub_printprocmem("complete load xlsx or csv data file and dictionary",logger)
        # Import dictionary
        s_dict_column =[AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL,"REQ","EXPLAINATION"]
        df_dict_micro = AL.readxlsorcsv_noheader(path_input,"dictionary_for_microbiology_data",s_dict_column,logger)
        df_dict_micro.loc[df_dict_micro[s_dict_column[1]].isnull() == True, s_dict_column[1]] = AC.CONST_DICVAL_EMPTY
        df_dict_hosp = pd.DataFrame()
        if bishosp_ava:
            df_dict_hosp = AL.readxlsorcsv_noheader(path_input,"dictionary_for_hospital_admission_data",s_dict_column,logger)
            df_dict_hosp.loc[df_dict_hosp[s_dict_column[1]].isnull() == True, s_dict_column[1]] = AC.CONST_DICVAL_EMPTY
        # The following part may be optional  --------------------------------------------------------------------------------------------------
        # Export the list of variables
        if not os.path.exists(path_variable):
            os.mkdir(path_variable)
        list_micro_val = list(df_micro.columns)
        df_micro_val = pd.DataFrame(list_micro_val)
        df_micro_val.columns = ['variables_micro']
        df_micro_val.to_csv(path_variable + "variables_micro.csv",index=False)
        if bishosp_ava:
            df_hosp_val = pd.DataFrame(list(df_hosp.columns))
            df_hosp_val.columns = ['variables_hosp']
            df_hosp_val.to_csv(path_variable + "variables_hosp.csv",index=False)
        # --------------------------------------------------------------------------------------------------------------------------------------
        #Combine data dict
        df_dict = pd.DataFrame()
        if bishosp_ava:
            df_combine_dict = [df_dict_micro.iloc[1:,0:2],df_dict_hosp.iloc[1:,0:2]]
            df_dict = pd.concat(df_combine_dict)
        else:
            df_dict = df_dict_micro.iloc[1:,0:2] 
        AL.printlog("Complete load dictionary file",False,logger)  
        #Data validation log for df_dict duplicated value
        try:
            temp_df = df_dict[df_dict[AC.CONST_DICTCOL_DATAVAL].str.strip() != ""]
            temp_df = temp_df.groupby([AC.CONST_DICTCOL_DATAVAL]).size().reset_index(name='counts')
            temp_df = temp_df[temp_df['counts'] > 1]
            temp_df =temp_df.merge(df_dict, how="inner", left_on=AC.CONST_DICTCOL_DATAVAL, right_on=AC.CONST_DICTCOL_DATAVAL,suffixes=("", "_NOTUSED"))
            list_notconsiderdup = ['no','No','NO','yes','Yes','YES','xxx_Can be changed in the dictionary_of_variable_data.csv_xxx',
                                   'Data values of the variable recorded for "organism" in your microbiology data file',
                                   'Data values of the variable recorded in your microbiology data file']
            temp_df = temp_df[~temp_df[AC.CONST_DICTCOL_DATAVAL].isin(list_notconsiderdup)]
            if len(temp_df) > 0:
                temp_df = temp_df[[AC.CONST_DICTCOL_DATAVAL,AC.CONST_DICTCOL_AMASS]]
            else:
                temp_df = pd.DataFrame(columns=[AC.CONST_DICTCOL_DATAVAL,AC.CONST_DICTCOL_AMASS])
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_dup.xlsx", logger):
                AL.printlog("Warning : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_dup.xlsx",False,logger)  
            del temp_df
        except Exception as e: 
            AL.printlog("Warning : check for ducplicate value in both dictionary: " +  str(e),False,logger)   
        
        AL.printlog("Complete check and log dictionary file variables",False,logger)  
        #Reverse order of dict -> to be the same as R when map
        df_dict = df_dict[::-1].reset_index(drop = True) 
        bisabom = True
        try:
            if df_dict[df_dict[AC.CONST_DICTCOL_AMASS] == "acinetobacter_spp_or_baumannii"].iloc[0][AC.CONST_DICTCOL_DATAVAL] != "organism_acinetobacter_baumannii":
                bisabom = False
        except Exception as e: 
            AL.printlog("Warning : unable to read acinetobacter_spp_or_baumannii configuration from dictionary: " +  str(e),False,logger)   
        bisentspp = True
        try:
            if df_dict[df_dict[AC.CONST_DICTCOL_AMASS] == "enterococcus_spp_or_faecalis_and_faecium"].iloc[0][AC.CONST_DICTCOL_DATAVAL] != "organism_enterococcus_spp":
                bisentspp = False
        except Exception as e: 
            AL.printlog("Warning : unable to read enterococcus_spp_or_faecalis_and_faecium configuration from dictionary: " +  str(e),False,logger)  
        # Grep function in R - may be the following Python code do it wrongly --- this may need to remove by changing the data dictionary
        if bisabom == False:
            df_dict.loc[df_dict[AC.CONST_DICTCOL_AMASS].str.contains("organism_acinetobacter"),AC.CONST_DICTCOL_AMASS] = "organism_acinetobacter_spp"
        if bisentspp == True:
            df_dict.loc[df_dict[AC.CONST_DICTCOL_AMASS].str.contains("organism_enterococcus"),AC.CONST_DICTCOL_AMASS] = "organism_enterococcus_spp"
        df_dict.loc[df_dict[AC.CONST_DICTCOL_AMASS].str.contains("organism_salmonella"),AC.CONST_DICTCOL_AMASS] = "organism_salmonella_spp"
        #lst_dict_amassvar = df_dict.iloc[0:,0:1]
        #lst_dict_dataval = df_dict.iloc[0:,1:2]
        dict_datavaltoamass = pd.Series(df_dict[AC.CONST_DICTCOL_AMASS].values,index=df_dict[AC.CONST_DICTCOL_DATAVAL].str.strip()).to_dict()
        dict_amasstodataval = pd.Series(df_dict[AC.CONST_DICTCOL_DATAVAL].values,index=df_dict[AC.CONST_DICTCOL_AMASS].str.strip()).to_dict()
        #dict for value replacement in data column 
        temp_df = df_dict[df_dict[AC.CONST_DICTCOL_AMASS].isin(["specimen_blood","specimen_cerebrospinal_fluid","specimen_genital_swab","specimen_respiratory_tract","specimen_stool","specimen_urine","specimen_others"])]
        dict_spectype_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip()).to_dict()
        #dict_spectype_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string")).to_dict()
        temp_df = df_dict[df_dict[AC.CONST_DICTCOL_AMASS].isin(["male","female"])]
        dict_gender_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip()).to_dict()
        #dict_gender_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string")).to_dict()
        temp_df = df_dict[df_dict[AC.CONST_DICTCOL_AMASS].isin(["community_origin","unknown_origin","hospital_origin"])]
        dict_inforg_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip()).to_dict()
        #dict_inforg_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string")).to_dict()
        temp_df = df_dict[df_dict[AC.CONST_DICTCOL_AMASS] == "died"]
        dict_died_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip()).to_dict()
        #dict_died_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string")).to_dict()
        #------------------------------------------------------------------------------------------------------------------------------------------------------
        # dictionary from AMASS_amr_const
        dict_ris = AC.dict_ris(df_dict)
        dict_ast = AC.dict_ast()
        dict_orgcatwithatb = AC.dict_orgcatwithatb(bisabom,bisentspp)
        dict_orgwithatb_mortality = AC.dict_orgwithatb_mortality(bisabom)
        dict_orgwithatb_incidence = AC.dict_orgwithatb_incidence(bisabom)
        list_antibiotic = AC.list_antibiotic
        AL.printlog("Complete process load data and dictionary: " + str(datetime.now()),False,logger)    
    except Exception as e: # work on python 3.x
        AL.printlog("Fail precoess load data and dictionary: " +  str(e),True,logger)
        logger.exception(e)
    #########################################################################################################################################
    # Start transform hosp and micro data
    try:
        # rename variable in dataval to amassval
        # What happen if key is duplicate
        df_micro.rename(columns=dict_datavaltoamass, inplace=True)   
        if bishosp_ava:
            df_hosp.rename(columns=dict_datavaltoamass, inplace=True)
        #Remove unused column to save memory
        list_micorcol = [AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_VARNAME_SPECDATERAW,AC.CONST_VARNAME_COHO,AC.CONST_VARNAME_ORG,AC.CONST_VARNAME_SPECTYPE]
        list_micorcol = list_micorcol + list_antibiotic
        df_micro = AL.fn_keeponlycol(df_micro, list_micorcol)
        #-----------------------------------------------------------------------------------------------------------------------------------------    
        # Assign row number as unique key for each microbiology data row (This is not in R code)
        df_micro[AC.CONST_NEWVARNAME_MICROREC_ID] = np.arange(len(df_micro))
        #-----------------------------------------------------------------------------------------------------------------------------------------
        # Trim of space and unreadable charector for field that may need to map values such as spectype, organism.
        df_micro = AL.fn_df_strstrips(df_micro,[AC.CONST_VARNAME_SPECTYPE,AC.CONST_VARNAME_ORG,AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_VARNAME_SPECNUM,AC.CONST_VARNAME_DISCHARGESTATUS,AC.CONST_VARNAME_GENDER,AC.CONST_VARNAME_COHO],logger)
        # Transform and map data
        #df_micro[AC.CONST_NEWVARNAME_BLOOD] = df_micro[AC.CONST_VARNAME_SPECTYPE].astype("string").str.strip().map(dict_spectype_datavaltoamass)
        df_micro[AC.CONST_NEWVARNAME_AMASSSPECTYPE] = df_micro[AC.CONST_VARNAME_SPECTYPE].astype("string").map(dict_spectype_datavaltoamass)
        #df_micro[AC.CONST_NEWVARNAME_BLOOD] = df_micro[AC.CONST_VARNAME_SPECTYPE].astype("string").map(dict_spectype_datavaltoamass)
        df_micro[AC.CONST_NEWVARNAME_BLOOD]  = df_micro[AC.CONST_NEWVARNAME_AMASSSPECTYPE]
        #df_micro.loc[df_micro[AC.CONST_NEWVARNAME_BLOOD] == "specimen_blood", "blood"] = "blood"
        #df_micro.loc[df_micro[AC.CONST_NEWVARNAME_BLOOD] != "blood", "blood"] = "non-blood"
        df_micro.loc[df_micro[AC.CONST_NEWVARNAME_BLOOD] == "specimen_blood", AC.CONST_NEWVARNAME_BLOOD] = "blood"
        df_micro.loc[df_micro[AC.CONST_NEWVARNAME_BLOOD] != "blood", AC.CONST_NEWVARNAME_BLOOD] = "non-blood"
        df_micro[AC.CONST_NEWVARNAME_AMASSSPECTYPE] = df_micro[AC.CONST_NEWVARNAME_AMASSSPECTYPE].fillna("undefined")
        """
        temp_df = df_micro.groupby([AC.CONST_NEWVARNAME_AMASSSPECTYPE]).size().reset_index(name='counts')
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_amassspectype.xlsx", logger):
            AL.printlog("Warning : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_amassspectype.xlsx",False,logger)  
        del temp_df
        """
        #df_micro[AC.CONST_NEWVARNAME_ORG3] = df_micro[AC.CONST_VARNAME_ORG].str.strip().map(dict_datavaltoamass).fillna(df_micro[AC.CONST_VARNAME_ORG])
        df_micro[AC.CONST_NEWVARNAME_ORG3] = df_micro[AC.CONST_VARNAME_ORG].astype("string").map(dict_datavaltoamass).fillna(df_micro[AC.CONST_VARNAME_ORG])
        df_micro = clean_hn(df_micro,AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_NEWVARNAME_HN)
        df_micro = fn_clean_date_andcalday_year_month(df_micro, AC.CONST_VARNAME_SPECDATERAW, AC.CONST_NEWVARNAME_CLEANSPECDATE, AC.CONST_NEWVARNAME_DAYTOSPECDATE, AC.CONST_NEWVARNAME_SPECYEAR, AC.CONST_NEWVARNAME_SPECMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
        df_micro = fn_clean_date_andcalday_year_month(df_micro, AC.CONST_VARNAME_SPECRPTDATERAW, AC.CONST_NEWVARNAME_CLEANSPECRPTDATE, AC.CONST_NEWVARNAME_DAYTOSPECRPTDATE, AC.CONST_NEWVARNAME_SPECRPTYEAR, AC.CONST_NEWVARNAME_SPECRPTMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
        dict_progvar["date_include_min"] = df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].min().strftime("%d %b %Y")
        dict_progvar["date_include_max"] = df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].max().strftime("%d %b %Y")
        # Transform data hm
        if bishosp_ava:
            fn_clean_ward(df_hosp,AC.CONST_NEWVARNAME_WARDCODE,AC.CONST_NEWVARNAME_WARDTYPE,path_input,"dictionary_for_wards",logger) 
            df_hosp = AL.fn_keeponlycol(df_hosp, [AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_VARNAME_ADMISSIONDATE,AC.CONST_VARNAME_DISCHARGEDATE,
                                                  AC.CONST_VARNAME_DISCHARGESTATUS,AC.CONST_VARNAME_GENDER,AC.CONST_VARNAME_BIRTHDAY,AC.CONST_VARNAME_AGEY,AC.CONST_VARNAME_AGEGROUP,
                                                  AC.CONST_VARNAME_WARD,AC.CONST_NEWVARNAME_WARDCODE,AC.CONST_NEWVARNAME_WARDTYPE])
            # Trim of space and unreadable charector for field that may need to map values such as spectype, organism.
            df_hosp = AL.fn_df_strstrips(df_hosp,[AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_VARNAME_DISCHARGESTATUS,AC.CONST_VARNAME_GENDER,AC.CONST_VARNAME_COHO],logger)
            df_hosp = clean_hn(df_hosp,AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_NEWVARNAME_HN_HOSP)
            # Convert field to category type to save mem space
            df_hosp = AL.fn_df_tocategory_datatype(df_hosp,[AC.CONST_VARNAME_DISCHARGESTATUS,AC.CONST_VARNAME_GENDER,AC.CONST_VARNAME_AGEGROUP,AC.CONST_VARNAME_COHO],logger)
            df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = 0
            if dict_amasstodataval['age_year_available'].lower() == "yes":
                df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = df_hosp[AC.CONST_VARNAME_AGEY].apply(pd.to_numeric,errors='coerce')
            elif dict_amasstodataval['birthday_available'].lower() == "yes":
                df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] =  (df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE] - df_hosp[AC.CONST_NEWVARNAME_DAYTOBIRTHDATE])/365.25
            df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = df_hosp[AC.CONST_NEWVARNAME_AGEYEAR].apply(np.floor,errors='coerce')
            df_hosp_formerge = df_hosp.copy(deep=True)
            #df_hosp[AC.CONST_NEWVARNAME_DISOUTCOME_HOSP] = df_hosp[AC.CONST_VARNAME_DISCHARGESTATUS].astype("string").str.strip().map(dict_died_datavaltoamass).fillna(AC.CONST_ALIVE_VALUE) # From R code line 1154
            df_hosp[AC.CONST_NEWVARNAME_DISOUTCOME_HOSP] = df_hosp[AC.CONST_VARNAME_DISCHARGESTATUS].astype("string").map(dict_died_datavaltoamass).fillna(AC.CONST_ALIVE_VALUE) # From R code line 1154
            df_hosp = fn_clean_date_andcalday_year_month(df_hosp, AC.CONST_VARNAME_ADMISSIONDATE, AC.CONST_NEWVARNAME_CLEANADMDATE, AC.CONST_NEWVARNAME_DAYTOADMDATE, AC.CONST_NEWVARNAME_ADMYEAR, AC.CONST_NEWVARNAME_ADMMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            df_hosp = fn_clean_date_andcalday_year_month(df_hosp, AC.CONST_VARNAME_DISCHARGEDATE, AC.CONST_NEWVARNAME_CLEANDISDATE, AC.CONST_NEWVARNAME_DAYTODISDATE, AC.CONST_NEWVARNAME_DISYEAR, AC.CONST_NEWVARNAME_DISMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            df_hosp = fn_clean_date_andcalday_year_month(df_hosp, AC.CONST_VARNAME_BIRTHDAY, AC.CONST_NEWVARNAME_CLEANBIRTHDATE, AC.CONST_NEWVARNAME_DAYTOBIRTHDATE, "","", AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            #V3.0.3 Calculate max date include
            df_hosp[AC.CONST_NEWVARNAME_DAYTOSTARTDATE] = df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE]
            df_hosp[AC.CONST_NEWVARNAME_DAYTOENDDATE] = df_hosp[AC.CONST_NEWVARNAME_DAYTODISDATE]
            try:
                dict_progvar["date_include_min"] = calcolmindayinclude(df_micro,AC.CONST_NEWVARNAME_CLEANSPECDATE,df_hosp,AC.CONST_NEWVARNAME_CLEANADMDATE,AC.CONST_NEWVARNAME_DAYTOSTARTDATE,AC.CONST_ORIGIN_DATE,dict_progvar["date_include_min"],logger)   
                dict_progvar["date_include_max"] = calcolmaxdayinclude(df_micro,AC.CONST_NEWVARNAME_CLEANSPECDATE,df_hosp,AC.CONST_NEWVARNAME_CLEANADMDATE,AC.CONST_NEWVARNAME_DAYTOENDDATE,AC.CONST_ORIGIN_DATE,dict_progvar["date_include_max"],logger) 
            except Exception as e:
                AL.printlog("Warning : calculate min max include date "  + str(e),True,logger)
                logger.exception(e)
                pass
            
            """
            if dtmp != "":
               dict_progvar["date_include_min"]  = dtmp
            if dtmp != "":
               dict_progvar["date_include_max"]  = dtmp
            """
            """
            try:
                print("--------------------------------------")
                dmax_data_include = df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].max()
                dtemp = df_hosp[AC.CONST_NEWVARNAME_CLEANADMDATE].max()
                if dtemp < dmax_data_include:
                    dmax_data_include = dtemp
                print(dmax_data_include)
                idaytomax_date_include = (dmax_data_include - AC.CONST_ORIGIN_DATE).days
                print(idaytomax_date_include)
            except Exception as e:
                AL.printlog("Warning : Fail to calculate end date data include: " +  str(e),False,logger)
                logger.exception(e)
            """
            # Patient day
            """
            df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY] = df_hosp[AC.CONST_NEWVARNAME_DAYTODISDATE] - df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE] + 1
            """
            df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY] = df_hosp[AC.CONST_NEWVARNAME_DAYTOENDDATE] - df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE] + 1
            df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY_HO] = df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY] - 2
            df_hosp.loc[df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY_HO] <0, AC.CONST_NEWVARNAME_PATIENTDAY_HO] = 0
            #df_hosp['ddd'] = df_hosp[AC.CONST_VARNAME_DISCHARGESTATUS].astype("string").str.decode("utf8",errors='coerce')
            #df_hosp['dddlu'] = df_hosp[AC.CONST_VARNAME_DISCHARGESTATUS].astype("string").str.encode('latin1',errors='ignore').str.decode("utf8",errors='ignore')
            
            debug_savecsv(df_hosp,path_repwithPID + "prepared_hospital_data.csv",bisdebug,1,logger)
        # Start Org Cat vs AST of interest -------------------------------------------------------------------------------------------------------
        # Suggest to be in a configuration files hide from user is better in term of coding
        df_micro["Temp" + AC.CONST_NEWVARNAME_ORG3] = df_micro[AC.CONST_NEWVARNAME_ORG3]
        df_micro.loc[df_micro[AC.CONST_NEWVARNAME_ORG3].str.strip() == "","Temp" + AC.CONST_NEWVARNAME_ORG3] = AC.CONST_ORG_NOGROWTH
        df_micro[AC.CONST_NEWVARNAME_ORGCAT] = df_micro["Temp" + AC.CONST_NEWVARNAME_ORG3].map({k : vs[0] for k, vs in dict_orgcatwithatb.items()}).fillna(0)
        df_micro.drop("Temp" + AC.CONST_NEWVARNAME_ORG3, axis=1, inplace=True)              
        # Gen RIS columns
        for satb in list_antibiotic:
            if satb in df_micro.columns:
                #df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = df_micro[satb].map(dict_ris).fillna("")
                df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb]  = ""
                for s_toreplace in dict_ris:
                    df_micro.loc[df_micro[satb].str.contains(s_toreplace), AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = dict_ris[s_toreplace]
                df_micro[satb] = df_micro[satb].astype("category")
            else:
                df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = "" 
            df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb].astype("category")
        # Gen AST columns
        for satb in list_antibiotic:
            df_micro[AC.CONST_NEWVARNAME_PREFIX_AST + satb] = df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb].map(dict_ast).fillna("NA") 
            df_micro[AC.CONST_NEWVARNAME_PREFIX_AST + satb] = df_micro[AC.CONST_NEWVARNAME_PREFIX_AST + satb].astype("category")
        # Antibiotic group
        """
        list_ast_asvalue = ["1","NA"]
        # Third generation cephalosporins
        temp_cond = [(df_micro["ASTCeftriaxone"]=="1") | (df_micro["ASTCefotaxime"]=="1") | (df_micro["ASTCeftazidime"]=="1") | (df_micro["ASTCefixime"]=="1"),
                     (df_micro["ASTCeftriaxone"]=="NA") & (df_micro["ASTCefotaxime"]=="NA") & (df_micro["ASTCeftazidime"]=="NA") & (df_micro["ASTCefixime"]=="NA")]
        df_micro[AC.CONST_NEWVARNAME_AST3GC] = np.select(temp_cond,list_ast_asvalue,default="0")
        # Carbapenem
        temp_cond = [(df_micro["ASTImipenem"]=="1") | (df_micro["ASTMeropenem"]=="1") | (df_micro["ASTErtapenem"]=="1") | (df_micro["ASTDoripenem"]=="1"),
                     (df_micro["ASTImipenem"]=="NA") & (df_micro["ASTMeropenem"]=="NA") & (df_micro["ASTErtapenem"]=="NA") & (df_micro["ASTDoripenem"]=="NA")]
        df_micro[AC.CONST_NEWVARNAME_ASTCBPN] = np.select(temp_cond,list_ast_asvalue,default="0")
        # Fluoroquinolones
        temp_cond = [(df_micro["ASTCiprofloxacin"]=="1") | (df_micro["ASTLevofloxacin"]=="1"),
                     (df_micro["ASTCiprofloxacin"]=="NA") & (df_micro["ASTLevofloxacin"]=="NA")]
        df_micro[AC.CONST_NEWVARNAME_ASTFRQ] = np.select(temp_cond,list_ast_asvalue,default="0")
        # Tetracyclines
        temp_cond = [(df_micro["ASTTigecycline"]=="1") | (df_micro["ASTMinocycline"]=="1"),
                     (df_micro["ASTTigecycline"]=="NA") & (df_micro["ASTMinocycline"]=="NA")]
        df_micro[AC.CONST_NEWVARNAME_ASTTETRA] = np.select(temp_cond,list_ast_asvalue,default="0")
        # Aminoglycosides
        temp_cond = [(df_micro["ASTGentamicin"]=="1") | (df_micro["ASTAmikacin"]=="1"),
                     (df_micro["ASTGentamicin"]=="NA") & (df_micro["ASTAmikacin"]=="NA")]
        df_micro[AC.CONST_NEWVARNAME_ASTAMINOGLY] = np.select(temp_cond,list_ast_asvalue,default="0")
        # Methicillins
        temp_cond = [(df_micro["ASTMethicillin"]=="1") | (df_micro["ASTOxacillin"]=="1") | (df_micro["ASTCefoxitin"]=="1"),
                     (df_micro["ASTMethicillin"]=="NA") & (df_micro["ASTOxacillin"]=="NA") & (df_micro["ASTCefoxitin"]=="NA")]
        df_micro[AC.CONST_NEWVARNAME_ASTMRSA] = np.select(temp_cond,list_ast_asvalue,default="0")
        # Penicillins
        temp_cond = [(df_micro["ASTPenicillin_G"]=="1") | (df_micro["ASTOxacillin"]=="1"),
                     (df_micro["ASTPenicillin_G"]=="NA") & (df_micro["ASTOxacillin"]=="NA")]
        df_micro[AC.CONST_NEWVARNAME_ASTPEN] = np.select(temp_cond,list_ast_asvalue,default="0")
        list_ast_asvalue = ["1","2"]
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_AST3GC] == "0") & ((df_micro[AC.CONST_NEWVARNAME_ASTCBPN] != "1") | (df_micro[AC.CONST_NEWVARNAME_ASTCBPN].isna() == True)),
                     (df_micro[AC.CONST_NEWVARNAME_AST3GC] == "1") & ((df_micro[AC.CONST_NEWVARNAME_ASTCBPN] != "1") | (df_micro[AC.CONST_NEWVARNAME_ASTCBPN].isna() == True))]
        df_micro[AC.CONST_NEWVARNAME_AST3GCCBPN] = np.select(temp_cond,list_ast_asvalue,default="0")
        """
        list_ris_asvalue = ["R","I","S"]
        # Third generation cephalosporins
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Ceftriaxone"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Cefotaxime"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + "Ceftazidime"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + "Cefixime"]=="R"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Ceftriaxone"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Cefotaxime"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + "Ceftazidime"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + "Cefixime"]=="I"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Ceftriaxone"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Cefotaxime"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + "Ceftazidime"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + "Cefixime"]=="S")
                     ]
        df_micro[AC.CONST_NEWVARNAME_AST3GC_RIS] = np.select(temp_cond,list_ris_asvalue,default="")
        df_micro[AC.CONST_NEWVARNAME_AST3GC] = df_micro[AC.CONST_NEWVARNAME_AST3GC_RIS].map(dict_ast).fillna("NA") 
        df_micro[AC.CONST_NEWVARNAME_AST3GC_RIS] = df_micro[AC.CONST_NEWVARNAME_AST3GC_RIS].astype("category")
        df_micro[AC.CONST_NEWVARNAME_AST3GC] = df_micro[AC.CONST_NEWVARNAME_AST3GC].astype("category")
        # Carbapenem
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Imipenem"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Meropenem"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Ertapenem"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Doripenem"]=="R"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Imipenem"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Meropenem"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Ertapenem"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Doripenem"]=="I"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Imipenem"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Meropenem"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Ertapenem"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Doripenem"]=="S")
                     ]
        df_micro[AC.CONST_NEWVARNAME_ASTCBPN_RIS] = np.select(temp_cond,list_ris_asvalue,default="")
        df_micro[AC.CONST_NEWVARNAME_ASTCBPN] = df_micro[AC.CONST_NEWVARNAME_ASTCBPN_RIS].map(dict_ast).fillna("NA")
        df_micro[AC.CONST_NEWVARNAME_ASTCBPN_RIS] = df_micro[AC.CONST_NEWVARNAME_ASTCBPN_RIS].astype("category")
        df_micro[AC.CONST_NEWVARNAME_ASTCBPN] = df_micro[AC.CONST_NEWVARNAME_ASTCBPN].astype("category")
        # Fluoroquinolones
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Ciprofloxacin"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Levofloxacin"]=="R"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Ciprofloxacin"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Levofloxacin"]=="I"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Ciprofloxacin"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Levofloxacin"]=="S")
                     ]
        df_micro[AC.CONST_NEWVARNAME_ASTFRQ_RIS] = np.select(temp_cond,list_ris_asvalue,default="")
        df_micro[AC.CONST_NEWVARNAME_ASTFRQ] = df_micro[AC.CONST_NEWVARNAME_ASTFRQ_RIS].map(dict_ast).fillna("NA")
        df_micro[AC.CONST_NEWVARNAME_ASTFRQ_RIS] = df_micro[AC.CONST_NEWVARNAME_ASTFRQ_RIS].astype("category")
        df_micro[AC.CONST_NEWVARNAME_ASTFRQ] = df_micro[AC.CONST_NEWVARNAME_ASTFRQ].astype("category")
        # Tetracyclines
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Tigecycline"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Minocycline"]=="R"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Tigecycline"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Minocycline"]=="I"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Tigecycline"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Minocycline"]=="S")
                     ]
        df_micro[AC.CONST_NEWVARNAME_ASTTETRA_RIS] = np.select(temp_cond,list_ris_asvalue,default="")
        df_micro[AC.CONST_NEWVARNAME_ASTTETRA] =  df_micro[AC.CONST_NEWVARNAME_ASTTETRA_RIS].map(dict_ast).fillna("NA")
        df_micro[AC.CONST_NEWVARNAME_ASTTETRA_RIS] = df_micro[AC.CONST_NEWVARNAME_ASTTETRA_RIS].astype("category")
        df_micro[AC.CONST_NEWVARNAME_ASTTETRA] = df_micro[AC.CONST_NEWVARNAME_ASTTETRA].astype("category")
        # Aminoglycosides
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Gentamicin"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Amikacin"]=="R"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Gentamicin"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Amikacin"]=="I"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Gentamicin"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Amikacin"]=="S")
                     ]
        df_micro[AC.CONST_NEWVARNAME_ASTAMINOGLY_RIS] = np.select(temp_cond,list_ris_asvalue,default="")
        df_micro[AC.CONST_NEWVARNAME_ASTAMINOGLY] = df_micro[AC.CONST_NEWVARNAME_ASTAMINOGLY_RIS].map(dict_ast).fillna("NA") 
        df_micro[AC.CONST_NEWVARNAME_ASTAMINOGLY_RIS] = df_micro[AC.CONST_NEWVARNAME_ASTAMINOGLY_RIS].astype("category")
        df_micro[AC.CONST_NEWVARNAME_ASTAMINOGLY] = df_micro[AC.CONST_NEWVARNAME_ASTAMINOGLY].astype("category")
        # Methicillins
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Methicillin"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Oxacillin"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Cefoxitin"]=="R"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Methicillin"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Oxacillin"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Cefoxitin"]=="I"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Methicillin"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Oxacillin"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Cefoxitin"]=="S")
                     ]
        df_micro[AC.CONST_NEWVARNAME_ASTMRSA_RIS] = np.select(temp_cond,list_ris_asvalue,default="")
        df_micro[AC.CONST_NEWVARNAME_ASTMRSA] = df_micro[AC.CONST_NEWVARNAME_ASTMRSA_RIS].map(dict_ast).fillna("NA") 
        df_micro[AC.CONST_NEWVARNAME_ASTMRSA_RIS] = df_micro[AC.CONST_NEWVARNAME_ASTMRSA_RIS].astype("category")
        df_micro[AC.CONST_NEWVARNAME_ASTMRSA] = df_micro[AC.CONST_NEWVARNAME_ASTMRSA].astype("category")
        # Penicillins
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Penicillin_G"]=="R") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Oxacillin"]=="R"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Penicillin_G"]=="I") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Oxacillin"]=="I"),
                     (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Penicillin_G"]=="S") | (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS +"Oxacillin"]=="S")
                     ]
        df_micro[AC.CONST_NEWVARNAME_ASTPEN_RIS] = np.select(temp_cond,list_ris_asvalue,default="")
        df_micro[AC.CONST_NEWVARNAME_ASTPEN] = df_micro[AC.CONST_NEWVARNAME_ASTPEN_RIS].map(dict_ast).fillna("NA") 
        df_micro[AC.CONST_NEWVARNAME_ASTPEN_RIS] = df_micro[AC.CONST_NEWVARNAME_ASTPEN_RIS].astype("category")
        df_micro[AC.CONST_NEWVARNAME_ASTPEN] = df_micro[AC.CONST_NEWVARNAME_ASTPEN].astype("category")
        #3GCCBPN
        list_ast_asvalue = ["1","2"]
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_AST3GC] == "0") & ((df_micro[AC.CONST_NEWVARNAME_ASTCBPN] != "1") | (df_micro[AC.CONST_NEWVARNAME_ASTCBPN].isna() == True)),
                     (df_micro[AC.CONST_NEWVARNAME_AST3GC] == "1") & ((df_micro[AC.CONST_NEWVARNAME_ASTCBPN] != "1") | (df_micro[AC.CONST_NEWVARNAME_ASTCBPN].isna() == True))]
        df_micro[AC.CONST_NEWVARNAME_AST3GCCBPN] = np.select(temp_cond,list_ast_asvalue,default="0")
        
        list_amr_atb = AC.getlist_amr_atb(dict_orgcatwithatb)
        if len(list_amr_atb) <=0:
            AL.printlog("Waring: list of antibiotic of interest for amr (get from dict_orgcatwithatb defined in AMASS_amr_const.py) is empty, application may pickup wrong record during deduplication by organism",False,logger)  
            df_micro[AC.CONST_NEWVARNAME_AMR] = 0
            df_micro[AC.CONST_NEWVARNAME_AMR_TESTED] = 0
        else:
            df_micro[AC.CONST_NEWVARNAME_AMR] = df_micro[list_amr_atb].apply(pd.to_numeric,errors='coerce').sum(axis=1, skipna=True)   
            df_micro[AC.CONST_NEWVARNAME_AMR_TESTED] = df_micro[list_amr_atb].apply(pd.to_numeric,errors='coerce').count(axis=1,numeric_only=True)   
        #Micro remove and alter column data type to save memory usage --------------------------------------------------------------------------------------
        #Change text/object field type to category type to save memory usage
        df_micro = AL.fn_df_tocategory_datatype(df_micro,
                                                [AC.CONST_NEWVARNAME_BLOOD,AC.CONST_NEWVARNAME_ORG3,AC.CONST_NEWVARNAME_ORGCAT,AC.CONST_VARNAME_SPECTYPE,AC.CONST_VARNAME_ORG,AC.CONST_VARNAME_COHO,AC.CONST_NEWVARNAME_AMASSSPECTYPE],
                                                logger)
        """
        df_micro[AC.CONST_NEWVARNAME_BLOOD] =df_micro[AC.CONST_NEWVARNAME_BLOOD].astype("category")
        df_micro[AC.CONST_NEWVARNAME_ORG3] = df_micro[AC.CONST_NEWVARNAME_ORG3].astype("category")
        df_micro[AC.CONST_NEWVARNAME_ORGCAT] = df_micro[AC.CONST_NEWVARNAME_ORGCAT].astype("category")
        #df_micro[AC.CONST_NEWVARNAME_AMRCAT] = df_micro[AC.CONST_NEWVARNAME_AMRCAT].astype("category")
        try:
            df_micro[AC.CONST_VARNAME_SPECTYPE] = df_micro[AC.CONST_VARNAME_SPECTYPE].astype("category")
            df_micro[AC.CONST_VARNAME_ORG] = df_micro[AC.CONST_VARNAME_ORG].astype("category")
            df_micro[AC.CONST_VARNAME_COHO] = df_micro[AC.CONST_VARNAME_COHO].astype("category")
        except Exception as e: # work on python 3.x
            AL.printlog("Warning, fail convert field type to category type: " +  str(e),False,logger) 
        """
        #Drop original antibiotic col to save memory usage
        #df_micro = AL.fn_removecol(df_micro,list_antibiotic) #using in ANNEX B don't remove
 
        ## Only blood specimen !!!!!!!!!!
        df_micro_blood = df_micro.loc[df_micro[AC.CONST_NEWVARNAME_BLOOD]=="blood"]
        ## Only interesting org cat !!!!!!!!!!
        df_micro_bsi = df_micro_blood.loc[df_micro_blood[AC.CONST_NEWVARNAME_ORGCAT]!=AC.CONST_ORG_NOTINTEREST_ORGCAT]
        debug_savecsv(df_micro,path_repwithPID + "prepared_microbiology.csv",bisdebug,1,logger)
        debug_savecsv(df_micro_blood,path_repwithPID + "prepared_microbiology_blood.csv",bisdebug,1,logger)
        debug_savecsv(df_micro_bsi,path_repwithPID + "prepared_microbiology_BSI.csv",bisdebug,1,logger)
        # May need improve performance in future !!!!!!!!
        #Version 3.0.3 Ward variables
        
        AL.printlog("Complete prepare data: " + str(datetime.now()),False,logger)        
    except Exception as e: # work on python 3.x
        AL.printlog("Fail prepare data: " +  str(e),True,logger)   
        logger.exception(e)
    sub_printprocmem("before start merge",logger)
    # Start Merge with hosp data -------------------------------------------------------------------------------------------------------------------------
    df_datalog_mergedlist = pd.DataFrame()
    try:            
        df_hospmicro = fn_mergededup_hospmicro(df_micro, df_hosp_formerge, bishosp_ava,df_dict,dict_datavaltoamass,dict_inforg_datavaltoamass,dict_gender_datavaltoamass,dict_died_datavaltoamass,logger,df_datalog_mergedlist)
        sub_printprocmem("finish merge micro and hosp data",logger)
        #Change to filter from df_hospmicro instead of call merge again
        df_hospmicro_blood = df_hospmicro.loc[df_hospmicro[AC.CONST_NEWVARNAME_BLOOD]=="blood"]
        sub_printprocmem("finish merge micro (blood) and hosp data",logger)
        df_hospmicro_bsi = df_hospmicro_blood.loc[df_hospmicro_blood[AC.CONST_NEWVARNAME_ORGCAT]!=AC.CONST_ORG_NOTINTEREST_ORGCAT]
        sub_printprocmem("finish merge micro (bsi) and hosp data",logger)
        """
        df_hospmicro_blood = fn_mergededup_hospmicro(df_micro_blood, df_hosp_formerge, bishosp_ava,df_dict,dict_datavaltoamass,dict_inforg_datavaltoamass,dict_gender_datavaltoamass,dict_died_datavaltoamass,logger)
        sub_printprocmem("finish merge micro (blood) and hosp data",logger)
        df_hospmicro_bsi = fn_mergededup_hospmicro(df_micro_bsi, df_hosp_formerge, bishosp_ava,df_dict,dict_datavaltoamass,dict_inforg_datavaltoamass,dict_gender_datavaltoamass,dict_died_datavaltoamass,logger)
        sub_printprocmem("finish merge micro (bsi) and hosp data",logger)
        """
        del df_hosp_formerge
        debug_savecsv(df_hospmicro,path_repwithPID + "merged_hospital_microbiology.csv",bisdebug,1,logger)
        debug_savecsv(df_hospmicro_blood,path_repwithPID + "merged_hospital_microbiology_blood.csv",bisdebug,1,logger)
        debug_savecsv(df_hospmicro_bsi,path_repwithPID + "merged_hospital_microbiology_BSI.csv",bisdebug,1,logger)
        gc.collect() 
        AL.printlog("Complete merge hosp and micro data: " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail merge hosp and micro data: " +  str(e),True,logger) 
        logger.exception(e)
    #########################################################################################################################################
    # Start summarize data - This part will not alter or add column to orignal micro data and merge hospital data 
    # ---------------------------------------------------------------------------------------------------------------------------------------
    # General summary variable for calculation/report
    try:
        if bishosp_ava:
            dict_progvar["patientdays"] =int(df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY].sum())
            dict_progvar["patientdays_ho"] =int(df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY_HO].sum())
            dict_progvar["hosp_date_min"] = df_hosp[AC.CONST_NEWVARNAME_CLEANADMDATE].min().strftime("%d %b %Y")
            dict_progvar["hosp_date_max"] = df_hosp[AC.CONST_NEWVARNAME_CLEANADMDATE].max().strftime("%d %b %Y")
        else:
            dict_progvar["patientdays"] ="NA"
            dict_progvar["patientdays_ho"] ="NA"
            dict_progvar["hosp_date_min"] = "NA"
            dict_progvar["hosp_date_max"] = "NA"
        #try:
        #except Exception as e: # work on python 3.x
        #AL.printlog("Warning analysis (By Sample base (Section 1,2,4,5,6)): " +  str(e),False,logger)    
        dict_progvar["micro_date_min"] = df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].min().strftime("%d %b %Y")
        dict_progvar["micro_date_max"] = df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].max().strftime("%d %b %Y")
        dict_progvar["n_blood"] = len(df_micro_blood)
        dict_progvar["n_blood_pos"] = len(df_micro_blood[df_micro_blood[AC.CONST_NEWVARNAME_ORGCAT] != AC.CONST_ORG_NOGROWTH_ORGCAT])
        dict_progvar["n_blood_neg"] = len(df_micro_blood[df_micro_blood[AC.CONST_NEWVARNAME_ORGCAT] == AC.CONST_ORG_NOGROWTH_ORGCAT])
        dict_progvar["n_bsi_pos"] = len(df_micro_bsi[df_micro_bsi[AC.CONST_NEWVARNAME_ORGCAT] != AC.CONST_ORG_NOGROWTH_ORGCAT])
        dict_progvar["n_blood_patients"] = len(df_micro_blood[AC.CONST_NEWVARNAME_HN].unique())
        dict_progvar["checkpoint_section6"] = 0
        try:
            dict_progvar["checkpoint_section6"] = len(df_hospmicro_bsi[(df_hospmicro_bsi[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_DIED_VALUE) | (df_hospmicro_bsi[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_ALIVE_VALUE)])
        except:
            dict_progvar["checkpoint_section6"] = 0
        df_isoRep_blood = pd.DataFrame(columns=["Organism","Antibiotic","Susceptible(N)","Non-susceptible(N)","Total(N)","Non-susceptible(%)","lower95CI(%)*","upper95CI(%)*",
                                                "Resistant(N)","Intermediate(N)",
                                                "Resistant(%)","Resistant-lower95CI(%)*","Resistant-upper95CI(%)*",
                                                "Intermediate(%)","Intermediate-lower95CI(%)*","Intermediate-upper95CI(%)*"])
        df_isoRep_blood_byorg = pd.DataFrame(columns=["Organism","Number_of_blood_specimens_culture_positive_for_the_organism"])
        df_isoRep_blood_byorg_dedup = pd.DataFrame(columns=["Organism","Number_of_blood_specimens_culture_positive_deduplicated"])
        df_isoRep_blood_incidence = pd.DataFrame(columns=["Organism","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci"])
        df_isoRep_blood_incidence_atb = pd.DataFrame(columns=["Organism","Priority_pathogen","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci"])
        for sorgkey in dict_orgcatwithatb:
            ocurorg = dict_orgcatwithatb[sorgkey]
            # skip no growth
            if ocurorg[1] == 1 :
                scurorgcat = ocurorg[0]
                temp_df = fn_deduplicatebyorgcat(df_micro_bsi,AC.CONST_NEWVARNAME_ORGCAT, int(scurorgcat))
                # Start for atb sensitivity summary
                sorgname = ocurorg[2]
                #onew_row = {"Antibiotics":scolname,"frequency_raw":str(n_row)}     
                #temp_df = pd.concat([temp_df,pd.DataFrame([onew_row])], ignore_index = True)
                
                
                #df_isoRep_blood_byorg = df_isoRep_blood_byorg.append({"Organism":sorgname,"Number_of_blood_specimens_culture_positive_for_the_organism":len(df_micro_bsi[df_micro_bsi[AC.CONST_NEWVARNAME_ORGCAT]==int(scurorgcat)])},ignore_index = True)
                df_isoRep_blood_byorg = pd.concat([df_isoRep_blood_byorg,pd.DataFrame([{"Organism":sorgname,"Number_of_blood_specimens_culture_positive_for_the_organism":len(df_micro_bsi[df_micro_bsi[AC.CONST_NEWVARNAME_ORGCAT]==int(scurorgcat)])}])],ignore_index = True)
                #df_isoRep_blood_byorg_dedup = df_isoRep_blood_byorg_dedup.append({"Organism":sorgname,"Number_of_blood_specimens_culture_positive_deduplicated":len(temp_df)},ignore_index = True)
                df_isoRep_blood_byorg_dedup = pd.concat([df_isoRep_blood_byorg_dedup,pd.DataFrame([{"Organism":sorgname,"Number_of_blood_specimens_culture_positive_deduplicated":len(temp_df)}])],ignore_index = True)
                list_atbname = ocurorg[3]
                list_atbmicrocol = ocurorg[4]
                # ADD CHECK if len list not the same avoid error
                for i in range(len(list_atbname)):
                    scuratbname = list_atbname[i]
                    scuratbmicrocol =list_atbmicrocol[i]
                    iRIS_S = 0
                    iRIS_I = 0
                    iRIS_R = 0
                    if scuratbmicrocol[0:len(AC.CONST_NEWVARNAME_PREFIX_RIS)] == AC.CONST_NEWVARNAME_PREFIX_RIS:
                        iRIS_R = len(temp_df[temp_df[scuratbmicrocol] == "R"])
                        iRIS_I = len(temp_df[temp_df[scuratbmicrocol] == "I"])
                        iRIS_S = len(temp_df[temp_df[scuratbmicrocol] == "S"])
                        
                    else:
                        iRIS_R = len(temp_df[temp_df[scuratbmicrocol] == "1"])
                        iRIS_S = len(temp_df[temp_df[scuratbmicrocol] == "0"])
                    itotal = iRIS_R +  iRIS_I + iRIS_S
                    iSuscep = iRIS_S
                    iNonsuscep = iRIS_R +  iRIS_I
                    nNonSuscepPercent = "NA"
                    nLowerCI = "NA"
                    nUpperCI = "NA"
                    #Version 3.0.2
                    nR_percent = "NA"
                    nI_percent = "NA"
                    nR_LowerCI = "NA"
                    nR_UpperCI = "NA"
                    nI_LowerCI = "NA"
                    nI_UpperCI = "NA"
                    if itotal != 0 :
                        nNonSuscepPercent = round(iNonsuscep / itotal, 2) * 100
                        nLowerCI = fn_wilson_lowerCI(x=iNonsuscep, n=itotal, conflevel=0.95, decimalplace=1)
                        nUpperCI  = fn_wilson_upperCI(x=iNonsuscep, n=itotal, conflevel=0.95, decimalplace=1)
                        if scuratbmicrocol[0:len(AC.CONST_NEWVARNAME_PREFIX_RIS)] == AC.CONST_NEWVARNAME_PREFIX_RIS:
                            nR_percent = round(iRIS_R / itotal, 2) * 100
                            nR_LowerCI = fn_wilson_lowerCI(x=iRIS_R, n=itotal, conflevel=0.95, decimalplace=1)
                            nR_UpperCI  = fn_wilson_upperCI(x=iRIS_R, n=itotal, conflevel=0.95, decimalplace=1)
                            nI_percent = round(iRIS_I / itotal, 2) * 100
                            nI_LowerCI = fn_wilson_lowerCI(x=iRIS_I, n=itotal, conflevel=0.95, decimalplace=1)
                            nI_UpperCI  = fn_wilson_upperCI(x=iRIS_I, n=itotal, conflevel=0.95, decimalplace=1)
                    onew_row = {"Organism":sorgname,"Antibiotic":scuratbname,"Susceptible(N)":iSuscep,"Non-susceptible(N)": iNonsuscep,"Total(N)":itotal,
                                "Non-susceptible(%)":nNonSuscepPercent,"lower95CI(%)*":nLowerCI,"upper95CI(%)*" : nUpperCI,
                                "Resistant(N)":iRIS_R,"Intermediate(N)":iRIS_I,
                                "Resistant(%)":nR_percent,"Resistant-lower95CI(%)*":nR_LowerCI,"Resistant-upper95CI(%)*":nR_UpperCI,
                                "Intermediate(%)":nI_percent,"Intermediate-lower95CI(%)*":nI_LowerCI,"Intermediate-upper95CI(%)*":nI_UpperCI}   
                    #df_isoRep_blood = df_isoRep_blood.append(onew_row,ignore_index = True)
                    df_isoRep_blood = pd.concat([df_isoRep_blood,pd.DataFrame([onew_row])],ignore_index = True)
                # Start incidence
                if sorgkey in dict_orgwithatb_incidence.keys():
                    ocurorg_incidence = dict_orgwithatb_incidence[sorgkey]
                    sorgname_incidence = ocurorg_incidence[0]
                    list_atbname_incidence = ocurorg_incidence[1]
                    list_atbmicrocol_incidence = ocurorg_incidence[2]
                    list_astvalue_incidence = ocurorg_incidence[3]
                    nPatient = len(temp_df)
                    nPercent = "NA"
                    nLowerCI = "NA"
                    nUpperCI = "NA"
                    if int(dict_progvar["n_blood_patients"]) > 0:
                        nPercent = (nPatient/dict_progvar["n_blood_patients"])*AC.CONST_PERPOP
                        nLowerCI = (AC.CONST_PERPOP/100)*fn_wilson_lowerCI(x=nPatient, n=dict_progvar["n_blood_patients"], conflevel=0.95, decimalplace=10)
                        nUpperCI  = (AC.CONST_PERPOP/100)*fn_wilson_upperCI(x=nPatient, n=dict_progvar["n_blood_patients"], conflevel=0.95, decimalplace=10)
                    onew_row = {"Organism":sorgname_incidence,"Number_of_patients":nPatient,"frequency_per_tested":nPercent,"frequency_per_tested_lci":nLowerCI,"frequency_per_tested_uci":nUpperCI  }   
                    #df_isoRep_blood_incidence = df_isoRep_blood_incidence.append(onew_row,ignore_index = True)
                    df_isoRep_blood_incidence = pd.concat([df_isoRep_blood_incidence,pd.DataFrame([onew_row])],ignore_index = True)
                    for i in range(len(list_atbname_incidence)):
                        scuratbname = list_atbname_incidence[i]
                        scuratbmicrocol =list_atbmicrocol_incidence[i]
                        sCurastvalue = list_astvalue_incidence[i]
                        nPatient = len(temp_df[temp_df[scuratbmicrocol] == sCurastvalue])
                        nPercent = "NA"
                        nLowerCI = "NA"
                        nUpperCI = "NA"
                        if int(dict_progvar["n_blood_patients"]) > 0:
                            nPercent = (nPatient/dict_progvar["n_blood_patients"])*AC.CONST_PERPOP
                            nLowerCI = (AC.CONST_PERPOP/100)*fn_wilson_lowerCI(x=nPatient, n=dict_progvar["n_blood_patients"], conflevel=0.95, decimalplace=10)
                            nUpperCI  = (AC.CONST_PERPOP/100)*fn_wilson_upperCI(x=nPatient, n=dict_progvar["n_blood_patients"], conflevel=0.95, decimalplace=10)
                        onew_row = {"Organism":sorgname_incidence,"Priority_pathogen":scuratbname,"Number_of_patients":nPatient,"frequency_per_tested":nPercent,"frequency_per_tested_lci":nLowerCI,"frequency_per_tested_uci":nUpperCI  }   
                        #df_isoRep_blood_incidence_atb = df_isoRep_blood_incidence_atb.append(onew_row,ignore_index = True)
                        df_isoRep_blood_incidence_atb = pd.concat([df_isoRep_blood_incidence_atb,pd.DataFrame([onew_row])],ignore_index = True)
        AL.printlog("Complete analysis (By Sample base (Section 1,2,4,5,6)): " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail analysis (By Sample base (Section 1,2,4,5,6)): " +  str(e),True,logger)     
        logger.exception(e)
    
    # --------------------------------------------------------------------------------------------------------------------------------------------------
    # Summary data from hospmicro_blood and hospmicro_bsi 
    # Separate CO/HO dataframe
    try:
        temp_df = df_hospmicro_blood.filter([AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_CLEANADMDATE, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_COHO_FINAL], axis=1)
        temp_df = fn_deduplicatedata(temp_df,[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_CLEANADMDATE, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True,True],"last",[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_CLEANADMDATE],"first")
        temp_df2 = temp_df.loc[temp_df[AC.CONST_NEWVARNAME_COHO_FINAL] == 0]
        temp_df = temp_df.loc[temp_df[AC.CONST_NEWVARNAME_COHO_FINAL] == 1]
        dict_progvar["n_CO_blood_patients"] = len(temp_df2[AC.CONST_NEWVARNAME_HN].unique())
        dict_progvar["n_HO_blood_patients"] = len(temp_df[AC.CONST_NEWVARNAME_HN].unique())
        temp_df = temp_df.merge(temp_df2, how="inner", left_on=AC.CONST_NEWVARNAME_HN, right_on=AC.CONST_NEWVARNAME_HN,suffixes=("", "CO"))
        dict_progvar["n_2adm_firstbothCOHO_patients"] = len(temp_df[AC.CONST_NEWVARNAME_HN].unique())
        df_COHO_isoRep_blood = pd.DataFrame(columns=["Organism","Infection_origin","Antibiotic","Susceptible(N)","Non-susceptible(N)","Total(N)","Non-susceptible(%)","lower95CI(%)*","upper95CI(%)*",
                                                     "Resistant(N)","Intermediate(N)",
                                                     "Resistant(%)","Resistant-lower95CI(%)*","Resistant-upper95CI(%)*",
                                                     "Intermediate(%)","Intermediate-lower95CI(%)*","Intermediate-upper95CI(%)*"])

        df_COHO_isoRep_blood_byorg = pd.DataFrame(columns=["Organism","Number_of_patients_with_blood_culture_positive","Number_of_patients_with_blood_culture_positive_merged_with_hospital_data_file","Community_origin","Hospital_origin","Unknown_origin"])
        df_COHO_isoRep_blood_mortality = pd.DataFrame(columns=["Organism","Infection_origin","Antibiotic","Mortality","Mortality_lower_95ci","Mortality_upper_95ci","Number_of_deaths","Total_number_of_patients"])
        df_COHO_isoRep_blood_mortality_byorg = pd.DataFrame(columns=["Organism","Infection_origin","Number_of_deaths","Total_number_of_patients"])
        df_COHO_isoRep_blood_incidence = pd.DataFrame(columns=["Organism","Infection_origin","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci"])
        df_COHO_isoRep_blood_incidence_atb = pd.DataFrame(columns=["Organism","Infection_origin","Priority_pathogen","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci"])
        for sorgkey in dict_orgcatwithatb:
            ocurorg = dict_orgcatwithatb[sorgkey]
            # skip no growth
            if ocurorg[1] == 1 :
                scurorgcat = ocurorg[0]
                temp_df_byorgcat = fn_deduplicatebyorgcat_hospmico(df_hospmicro_bsi,AC.CONST_NEWVARNAME_ORGCAT, int(scurorgcat))
                debug_savecsv(temp_df_byorgcat,path_repwithPID + "section3_dedup_"  + ocurorg[2] + ".csv",bisdebug,1,logger)
                iCO = 0
                iHO = 0
                for iCOHO in range(2):
                    # Dedup here may be differ from R code as R dedup before filter !!!
                    if iCOHO == 0:
                        sCOHO = AC.CONST_EXPORT_COHO_CO_DATAVAL
                        sCOHO_mortality = AC.CONST_EXPORT_COHO_MORTALITY_CO_DATAVAL
                        temp_df = temp_df_byorgcat.loc[temp_df_byorgcat[AC.CONST_NEWVARNAME_COHO_FINAL] == 0]
                        iCO = len(temp_df)
                    else:
                        sCOHO = AC.CONST_EXPORT_COHO_HO_DATAVAL
                        sCOHO_mortality = AC.CONST_EXPORT_COHO_MORTALITY_HO_DATAVAL
                        temp_df = temp_df_byorgcat.loc[temp_df_byorgcat[AC.CONST_NEWVARNAME_COHO_FINAL] == 1]
                        iHO = len(temp_df)
                    # Start for atb sensitivity summary                    
                    sorgname = ocurorg[2]
                    list_atbname = ocurorg[3]
                    list_atbmicrocol = ocurorg[4]
                    # ADD CHECK if len list not the same avoid error
                    for i in range(len(list_atbname)):
                        scuratbname = list_atbname[i]
                        scuratbmicrocol =list_atbmicrocol[i]
                        iRIS_S = 0
                        iRIS_I = 0
                        iRIS_R = 0
                        if scuratbmicrocol[0:len(AC.CONST_NEWVARNAME_PREFIX_RIS)] == AC.CONST_NEWVARNAME_PREFIX_RIS:
                            iRIS_R = len(temp_df[temp_df[scuratbmicrocol] == "R"])
                            iRIS_I = len(temp_df[temp_df[scuratbmicrocol] == "I"])
                            iRIS_S = len(temp_df[temp_df[scuratbmicrocol] == "S"])
                            
                        else:
                            iRIS_R = len(temp_df[temp_df[scuratbmicrocol] == "1"])
                            iRIS_S = len(temp_df[temp_df[scuratbmicrocol] == "0"])
                        itotal = iRIS_R +  iRIS_I + iRIS_S
                        iSuscep = iRIS_S
                        iNonsuscep = iRIS_R +  iRIS_I
                        nNonSuscepPercent = "NA"
                        nLowerCI = "NA"
                        nUpperCI = "NA"
                        #Version 3.0.2
                        nR_percent = "NA"
                        nI_percent = "NA"
                        nR_LowerCI = "NA"
                        nR_UpperCI = "NA"
                        nI_LowerCI = "NA"
                        nI_UpperCI = "NA"
                        if itotal != 0 :
                            nNonSuscepPercent = round(iNonsuscep / itotal, 2) * 100
                            nLowerCI = fn_wilson_lowerCI(x=iNonsuscep, n=itotal, conflevel=0.95, decimalplace=1)
                            nUpperCI  = fn_wilson_upperCI(x=iNonsuscep, n=itotal, conflevel=0.95, decimalplace=1)
                            if scuratbmicrocol[0:len(AC.CONST_NEWVARNAME_PREFIX_RIS)] == AC.CONST_NEWVARNAME_PREFIX_RIS:
                                nR_percent = round(iRIS_R / itotal, 2) * 100
                                nR_LowerCI = fn_wilson_lowerCI(x=iRIS_R, n=itotal, conflevel=0.95, decimalplace=1)
                                nR_UpperCI  = fn_wilson_upperCI(x=iRIS_R, n=itotal, conflevel=0.95, decimalplace=1)
                                nI_percent = round(iRIS_I / itotal, 2) * 100
                                nI_LowerCI = fn_wilson_lowerCI(x=iRIS_I, n=itotal, conflevel=0.95, decimalplace=1)
                                nI_UpperCI  = fn_wilson_upperCI(x=iRIS_I, n=itotal, conflevel=0.95, decimalplace=1)
                        onew_row = {"Organism":sorgname,"Infection_origin":sCOHO,"Antibiotic":scuratbname,"Susceptible(N)":iSuscep,"Non-susceptible(N)": iNonsuscep,"Total(N)":itotal,
                                    "Non-susceptible(%)":nNonSuscepPercent,"lower95CI(%)*":nLowerCI,"upper95CI(%)*" : nUpperCI,
                                    "Resistant(N)":iRIS_R,"Intermediate(N)":iRIS_I,
                                    "Resistant(%)":nR_percent,"Resistant-lower95CI(%)*":nR_LowerCI,"Resistant-upper95CI(%)*":nR_UpperCI,
                                    "Intermediate(%)":nI_percent,"Intermediate-lower95CI(%)*":nI_LowerCI,"Intermediate-upper95CI(%)*":nI_UpperCI }
                        #df_COHO_isoRep_blood = df_COHO_isoRep_blood.append(onew_row,ignore_index = True)          
                        df_COHO_isoRep_blood = pd.concat([df_COHO_isoRep_blood,pd.DataFrame([onew_row])],ignore_index = True)  
                    # Start mortality
                    if sorgkey in dict_orgwithatb_mortality.keys():
                        ocurorg_mortality = dict_orgwithatb_mortality[sorgkey]
                        sorgname_mortality = ocurorg_mortality[0]
                        iOrgDied = len(temp_df[temp_df[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_DIED_VALUE])
                        iOrgAlive = len(temp_df[temp_df[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_ALIVE_VALUE])
                        iOrgTotal = iOrgDied + iOrgAlive
                        list_atbname_mortality = ocurorg_mortality[1]
                        list_atbmicrocol_mortality = ocurorg_mortality[2]
                        list_astvalue_mortality = ocurorg_mortality[3]
                        # ADD CHECK if len list not the same avoid error
                        for i in range(len(list_atbname_mortality)):
                            scuratbname = list_atbname_mortality[i]
                            scuratbmicrocol =list_atbmicrocol_mortality[i]
                            sCurastvalue = list_astvalue_mortality[i]
                            temp_df2 = temp_df.loc[temp_df[scuratbmicrocol] == sCurastvalue]
                            iDied = len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_DIED_VALUE])
                            iAlive= len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_ALIVE_VALUE])
                            itotal = iDied + iAlive
                            nLowerCI = 0
                            nUpperCI = 0
                            nDiedPercent = 0
                            sMortality = "NA"
                            if itotal != 0 :
                                nLowerCI = fn_wilson_lowerCI(x=iDied, n=itotal, conflevel=0.95, decimalplace=1)
                                nUpperCI  = fn_wilson_upperCI(x=iDied, n=itotal, conflevel=0.95, decimalplace=1)
                                nDiedPercent = int(round((iDied/itotal)*100, 0))
                                sMortality = str(nDiedPercent) + "% " + "(" + str(iDied) + "/" + str(itotal) + ")"
                            onew_row = {"Organism":sorgname_mortality,"Infection_origin":sCOHO_mortality,"Antibiotic":scuratbname,"Mortality":sMortality,
                                        "Mortality_lower_95ci":nLowerCI,"Mortality_upper_95ci":nUpperCI,"Number_of_deaths":iDied,"Total_number_of_patients":itotal}   
                            #df_COHO_isoRep_blood_mortality = df_COHO_isoRep_blood_mortality.append(onew_row,ignore_index = True)
                            df_COHO_isoRep_blood_mortality = pd.concat([df_COHO_isoRep_blood_mortality,pd.DataFrame([onew_row])],ignore_index = True)         
                        #df_COHO_isoRep_blood_mortality_byorg = df_COHO_isoRep_blood_mortality_byorg.append({"Organism":sorgname_mortality,"Infection_origin":sCOHO_mortality,
                        #                                                                                    "Number_of_deaths":iOrgDied,"Total_number_of_patients":iOrgTotal},ignore_index = True)
                        df_COHO_isoRep_blood_mortality_byorg = pd.concat([df_COHO_isoRep_blood_mortality_byorg,pd.DataFrame([{"Organism":sorgname_mortality,"Infection_origin":sCOHO_mortality,
                                                                                                            "Number_of_deaths":iOrgDied,"Total_number_of_patients":iOrgTotal}])],ignore_index = True)
                    # Start incidence
                    if sorgkey in dict_orgwithatb_incidence.keys():
                        ocurorg_incidence = dict_orgwithatb_incidence[sorgkey]
                        sorgname_incidence = ocurorg_incidence[0]
                        list_atbname_incidence = ocurorg_incidence[1]
                        list_atbmicrocol_incidence = ocurorg_incidence[2]
                        list_astvalue_incidence = ocurorg_incidence[3]
                        nPatient = len(temp_df)
                        nPercent = "NA"
                        nLowerCI = "NA"
                        nUpperCI = "NA"
                        if iCOHO == 0: 
                            nTotal = dict_progvar["n_CO_blood_patients"]
                        else:
                            nTotal = dict_progvar["n_HO_blood_patients"]
                        if int(nTotal) > 0:
                            nPercent = (nPatient/nTotal)*AC.CONST_PERPOP
                            nLowerCI = (AC.CONST_PERPOP/100)*fn_wilson_lowerCI(x=nPatient, n=nTotal, conflevel=0.95, decimalplace=10)
                            nUpperCI  = (AC.CONST_PERPOP/100)*fn_wilson_upperCI(x=nPatient, n=nTotal, conflevel=0.95, decimalplace=10)
                        onew_row = {"Organism":sorgname_incidence,"Infection_origin":sCOHO,"Number_of_patients":nPatient,"frequency_per_tested":nPercent,"frequency_per_tested_lci":nLowerCI,"frequency_per_tested_uci":nUpperCI  }   
                        #df_COHO_isoRep_blood_incidence = df_COHO_isoRep_blood_incidence.append(onew_row,ignore_index = True)
                        df_COHO_isoRep_blood_incidence = pd.concat([df_COHO_isoRep_blood_incidence,pd.DataFrame([onew_row])],ignore_index = True)
                        for i in range(len(list_atbname_incidence)):
                            scuratbname = list_atbname_incidence[i]
                            scuratbmicrocol =list_atbmicrocol_incidence[i]
                            sCurastvalue = list_astvalue_incidence[i]
                            nPatient = len(temp_df[temp_df[scuratbmicrocol] == sCurastvalue])
                            nPercent = "NA"
                            nLowerCI = "NA"
                            nUpperCI = "NA"
                            if int(nTotal) > 0:
                                nPercent = (nPatient/nTotal)*AC.CONST_PERPOP
                                nLowerCI = (AC.CONST_PERPOP/100)*fn_wilson_lowerCI(x=nPatient, n=nTotal, conflevel=0.95, decimalplace=10)
                                nUpperCI  = (AC.CONST_PERPOP/100)*fn_wilson_upperCI(x=nPatient, n=nTotal, conflevel=0.95, decimalplace=10)
                            onew_row = {"Organism":sorgname_incidence,"Infection_origin":sCOHO,"Priority_pathogen":scuratbname,"Number_of_patients":nPatient,"frequency_per_tested":nPercent,"frequency_per_tested_lci":nLowerCI,"frequency_per_tested_uci":nUpperCI  }   
                            #df_COHO_isoRep_blood_incidence_atb = df_COHO_isoRep_blood_incidence_atb.append(onew_row,ignore_index = True)   
                            df_COHO_isoRep_blood_incidence_atb = pd.concat([df_COHO_isoRep_blood_incidence_atb,pd.DataFrame([onew_row])],ignore_index = True) 
                #Save count CO/HO by org
                temp_df = fn_deduplicatebyorgcat(df_micro_bsi,AC.CONST_NEWVARNAME_ORGCAT, int(scurorgcat))
                iAll = len(temp_df)
                """
                df_COHO_isoRep_blood_byorg = df_COHO_isoRep_blood_byorg.append({"Organism":sorgname,
                                                                                "Number_of_patients_with_blood_culture_positive":iAll,
                                                                                "Number_of_patients_with_blood_culture_positive_merged_with_hospital_data_file":iCO + iHO,
                                                                                "Community_origin":iCO,
                                                                                "Hospital_origin":iHO,
                                                                                "Unknown_origin":iAll- iCO- iHO
                                                                                },ignore_index = True)
                """
                df_COHO_isoRep_blood_byorg = pd.concat([df_COHO_isoRep_blood_byorg,pd.DataFrame([{"Organism":sorgname,
                                                                                "Number_of_patients_with_blood_culture_positive":iAll,
                                                                                "Number_of_patients_with_blood_culture_positive_merged_with_hospital_data_file":iCO + iHO,
                                                                                "Community_origin":iCO,
                                                                                "Hospital_origin":iHO,
                                                                                "Unknown_origin":iAll- iCO- iHO
                                                                                }])],ignore_index = True)
        AL.printlog("Complete analysis (By Infection Origin (Section 3, 5, 6)): " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail analysis (By Infection Origin (Section 3, 5, 6)): " +  str(e),True,logger)  
        logger.exception(e)
    sub_printprocmem("finish section 1-6",logger)
    # --------------------------------------------------------------------------------------------------------------------------------------------------
    # ANNEX A
    try:
        dict_annex_a_listorg = AC.dict_annex_a_listorg
        dict_annex_a_spectype = AC.dict_annex_a_spectype
        
        df_dict_annex_a = df_dict_micro.copy(deep=True)
        #Reverse order of dict -> to be the same as R when map
        df_dict_annex_a = df_dict_annex_a[::-1].reset_index(drop = True) 
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS].str.contains("organism_brucella"),AC.CONST_DICTCOL_AMASS] = "organism_brucella_spp"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS].str.contains("organism_shigella"),AC.CONST_DICTCOL_AMASS] = "organism_shigella_spp"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS].str.contains("organism_vibrio"),AC.CONST_DICTCOL_AMASS] = "organism_vibrio_spp"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS] == "organism_salmonella_typhi",AC.CONST_DICTCOL_AMASS] = "t_typhi"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS] == "organism_salmonella_paratyphi",AC.CONST_DICTCOL_AMASS] = "t_paratyphi"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS].str.contains("organism_salmonella"),AC.CONST_DICTCOL_AMASS] = "organism_non-typhoidal_salmonella_spp"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS] == "t_typhi",AC.CONST_DICTCOL_AMASS] = "organism_salmonella_typhi"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS] == "t_paratyphi",AC.CONST_DICTCOL_AMASS] = "organism_salmonella_paratyphi"
        dict_datavaltoamass_annex_a = pd.Series(df_dict_annex_a[AC.CONST_DICTCOL_AMASS].values,index=df_dict_annex_a[AC.CONST_DICTCOL_DATAVAL].str.strip()).to_dict()
        df_micro_annex_a = df_micro.copy(deep=True)
        df_micro_annex_a[AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA] =  df_micro_annex_a[AC.CONST_VARNAME_SPECTYPE].str.strip().map(dict_datavaltoamass_annex_a).fillna("specimen_others")
        df_micro_annex_a[AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA] = df_micro_annex_a[AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA].str.strip().map(dict_annex_a_spectype).fillna("Others")
        df_micro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA] =  df_micro_annex_a[AC.CONST_VARNAME_ORG].str.strip().map(dict_datavaltoamass_annex_a).fillna(AC.CONST_ANNEXA_NON_ORG)
        df_micro_annex_a[AC.CONST_NEWVARNAME_ORGCAT_ANNEXA] =  df_micro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA].map({k : vs[0] for k, vs in dict_annex_a_listorg.items()}).fillna(0)
        df_micro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA] =  df_micro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA].map({k : vs[2] for k, vs in dict_annex_a_listorg.items()}).fillna("")
        df_micro_annex_a = df_micro_annex_a.loc[df_micro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA]!=""]
        df_micro_annex_a = df_micro_annex_a.loc[df_micro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA]!=AC.CONST_ANNEXA_NON_ORG]
        # Start Annex A summary table -----------------------------------------------------------------------------------------------------------
        temp_list = [['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                     ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]], 
                     ['microbiology_data','Number_of_all_culture_positive', len(df_micro_annex_a)]
                     ]
        for sspeckey in dict_annex_a_spectype:
            sspecname = dict_annex_a_spectype[sspeckey]
            sparameter = "Number_of_" + sspecname.replace("\n"," ").replace(" ","_").lower().strip() + "_culture_positive"
            n_annexa_speccount = len(df_micro_annex_a[df_micro_annex_a[AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA]==sspecname])
            temp_list.append(['microbiology_data',sparameter,str(n_annexa_speccount)])
        df_annexA = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"])
        # Start Annex A1 summary table ----------------------------------------------------------------------------------------------------------
        temp_df = fn_deduplicatedata(df_micro_annex_a,[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True,True,True],"last",[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA],"first")
        temp_df2 = fn_deduplicatedata(df_micro_annex_a,[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True,True],"last",[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA],"first")

        df_annexA1_pivot = temp_df.pivot_table(index=AC.CONST_NEWVARNAME_ORGNAME_ANNEXA, columns=AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA, aggfunc={AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA:len}, fill_value=0)
        #df_annexA1_pivot = temp_df.pivot_table(index=AC.CONST_NEWVARNAME_ORG3_ANNEXA, columns=AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA, aggfunc={AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA:len}, fill_value=0)

        df_annexA1_pivot.columns = df_annexA1_pivot.columns.droplevel(0)
        # add missing column and row to pivot table
        df_annexA1 = pd.DataFrame(columns=["Organism","Total number\nof patients*"])
        for oorg in dict_annex_a_listorg.values():
            if oorg[1] == 1 :
                sorg = oorg[2]
                #sorg = oorg[0]
                list_rowvalue = [sorg]
                #itotal = 0
                for sspec in dict_annex_a_spectype.values():
                    # add new column on the fly
                    if sspec not in df_annexA1.columns:
                        df_annexA1[sspec] = 0
                    # if have value for this org and specimen type then use it value
                    iCur = 0
                    if sorg in df_annexA1_pivot.index:
                        if sspec in df_annexA1_pivot.columns:
                            iCur = df_annexA1_pivot.loc[sorg][sspec]
                    list_rowvalue.append(iCur)
                    #itotal = itotal + iCur
                # append row  
                itotal = len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA]==sorg])
                #itotal = len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_ORG3_ANNEXA]==sorg])
                list_rowvalue.insert(1,itotal)
                df_annexA1.loc[sorg] = list_rowvalue 
        df_annexA1.loc['Total'] = df_annexA1.sum()
        df_annexA1.loc['Total','Organism'] = "Total"
        sub_printprocmem("finish analyse annex A1 data",logger)
        # Start Annex A2 summary table ----------------------------------------------------------------------------------------------------------
        #df_hospmicro_annex_a = df_hospmicro.copy(deep=True) - No need deep copy only using in ANNEX A
        df_hospmicro_annex_a = df_hospmicro
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA] =  df_hospmicro_annex_a[AC.CONST_VARNAME_SPECTYPE].str.strip().map(dict_datavaltoamass_annex_a).fillna("specimen_others")
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA] = df_hospmicro_annex_a[AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA].str.strip().map(dict_annex_a_spectype).fillna("Others")
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA] =  df_hospmicro_annex_a[AC.CONST_VARNAME_ORG].str.strip().map(dict_datavaltoamass_annex_a).fillna(AC.CONST_ANNEXA_NON_ORG)
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORGCAT_ANNEXA] =  df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA].map({k : vs[0] for k, vs in dict_annex_a_listorg.items()}).fillna(0)
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA] =  df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA].map({k : vs[2] for k, vs in dict_annex_a_listorg.items()}).fillna("")
        df_hospmicro_annex_a = df_hospmicro_annex_a.loc[df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA]!=""]
        df_hospmicro_annex_a = df_hospmicro_annex_a.loc[df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA]!=AC.CONST_ANNEXA_NON_ORG]
        temp_df = fn_deduplicatedata(df_hospmicro_annex_a,[AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True],"last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
        df_annexA2 = pd.DataFrame(columns=["Organism","Number_of_deaths","Total_number_of_patients","Mortality(%)","Mortality_lower_95ci","Mortality_upper_95ci"])
        for sorgkey in dict_annex_a_listorg:
            ocurorg = dict_annex_a_listorg[sorgkey]
            if ocurorg[1] == 1 :
                scurorgcat = ocurorg[0]
                sorgname = ocurorg[2]
                temp_df2 = temp_df.loc[temp_df[AC.CONST_NEWVARNAME_ORG3_ANNEXA]==sorgkey]
                iDied = len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_DIED_VALUE])
                iAlive= len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_ALIVE_VALUE])
                itotal = iDied + iAlive
                nLowerCI = 0
                nUpperCI = 0
                nDiedPercent = 0
                sMortality = "NA"
                if itotal != 0 :
                    nLowerCI = fn_wilson_lowerCI(x=iDied, n=itotal, conflevel=0.95, decimalplace=1)
                    nUpperCI  = fn_wilson_upperCI(x=iDied, n=itotal, conflevel=0.95, decimalplace=1)
                    nDiedPercent = int(round((iDied/itotal)*100, 0))
                    #sMortality = str(nDiedPercent) + "% " + "(" + str(iDied) + "/" + str(itotal) + ")"
                onew_row = {"Organism":sorgname,"Number_of_deaths":iDied,"Total_number_of_patients":itotal,
                            "Mortality(%)":nDiedPercent,"Mortality_lower_95ci":nLowerCI,"Mortality_upper_95ci":nUpperCI }   
                #df_annexA2 = df_annexA2.append(onew_row,ignore_index = True)  
                df_annexA2 = pd.concat([df_annexA2,pd.DataFrame([onew_row])],ignore_index = True)
        sub_printprocmem("finish analyse annex A2 data",logger)
        AL.printlog("Complete analysis (Annex A): " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail analysis (By Infection Origin (Section 3, 5, 6)): " +  str(e),True,logger) 
        logger.exception(e)        
    
    #########################################################################################################################################
    # Start export data - in future may also directly call function for generate report here by passing dataframe for better performance
    # --------------------------------------------------------------------------------------------------------------------------------------------------
    try:
        # --------------------------------------------------------------------------------------------------------------------------------------------------             
        # SECTION 1 - Number
        df_report1_page3 = pd.DataFrame(columns=["Type_of_data_file","Parameters","Values"])
        #temp_var_list = ["country","hospital_name","contact_person","contact_address","contact_email","notes_on_the_cover"]
        temp_var_list = [["hospital_name","Hospital_name"],["country","Country"],["contact_person","Contact_person"],["contact_address","Contact_address"],["contact_email","Contact_email"],["notes_on_the_cover","notes_on_the_cover"]]
        for ovar in temp_var_list:
            temp_list = ["microbiology_data", ovar[1],  dict_amasstodataval[ovar[0]]]
            df_report1_page3.loc[len(df_report1_page3)] = temp_list
        if bishosp_ava:
            temp_list = [['overall_data','Minimum_date', dict_progvar["date_include_min"]], 
                         ['overall_data','Maximum_date', dict_progvar["date_include_max"]], 
                         ['microbiology_data','Number_of_records', len(df_micro)], 
                         ['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                         ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]], 
                         ['hospital_admission_data','Number_of_records', len(df_hosp)], 
                         ['hospital_admission_data','Minimum_date', dict_progvar["hosp_date_min"]], 
                         ['hospital_admission_data','Maximum_date', dict_progvar["hosp_date_max"]],
                         ['hospital_admission_data','Patient_days', dict_progvar["patientdays"]],
                         ['hospital_admission_data','Patient_days_his', dict_progvar["patientdays_ho"]]]     
        else:
            temp_list = [['microbiology_data','Number_of_records', len(df_micro)], 
                         ['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                         ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]]]
        temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
        #df_report1_page3 = df_report1_page3.append(temp_df,ignore_index = True)
        df_report1_page3 = pd.concat([df_report1_page3,temp_df],ignore_index = True)
        if not AL.fn_savecsv(df_report1_page3, AC.CONST_PATH_RESULT + "Report1_page3_results.csv", 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report1_page3_results.csv")
        # SECTION 1 - by month
        df_report1_page4 = pd.DataFrame(data={"Month":fn_allmonthname()},index=fn_allmonthname())
        temp_df2 = df_micro.groupby([AC.CONST_NEWVARNAME_SPECMONTHNAME])[AC.CONST_NEWVARNAME_HN].count().reset_index(name='Number_of_specimen_in_microbiology_data_file')
        df_report1_page4 = df_report1_page4.merge(temp_df2, how="left", left_on="Month", right_on=AC.CONST_NEWVARNAME_SPECMONTHNAME,suffixes=("", "_MICRO"))
        df_report1_page4.drop(AC.CONST_NEWVARNAME_SPECMONTHNAME, axis=1, inplace=True)
        df_report1_page4['Number_of_specimen_in_microbiology_data_file'] = df_report1_page4['Number_of_specimen_in_microbiology_data_file'].fillna(0)
        if bishosp_ava:
            temp_df2 = df_hosp.groupby([AC.CONST_NEWVARNAME_ADMMONTHNAME])[AC.CONST_NEWVARNAME_HN_HOSP].count().reset_index(name='Number_of_hospital_records_in_hospital_admission_data_file')
            df_report1_page4 = df_report1_page4.merge(temp_df2, how="left", left_on="Month", right_on=AC.CONST_NEWVARNAME_ADMMONTHNAME,suffixes=("", "_HOSP"))
            df_report1_page4.drop(AC.CONST_NEWVARNAME_ADMMONTHNAME, axis=1, inplace=True)
            df_report1_page4['Number_of_hospital_records_in_hospital_admission_data_file'] = df_report1_page4['Number_of_hospital_records_in_hospital_admission_data_file'].fillna(0)
        if not AL.fn_savecsv(df_report1_page4, AC.CONST_PATH_RESULT + "Report1_page4_counts_by_month.csv", 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report1_page4_counts_by_month.csv")
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 2 - Summary
        temp_list = [['microbiology_data','Number_of_blood_specimens_collected', dict_progvar["n_blood"]], 
                     ['microbiology_data','Number_of_blood_culture_negative', dict_progvar["n_blood_neg"]], 
                     ['microbiology_data','Number_of_blood_culture_positive', dict_progvar["n_blood_pos"]], 
                     ['microbiology_data','Number_of_blood_culture_positive_for_organism_under_this_survey', dict_progvar["n_bsi_pos"]], 
                     ['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                     ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]]]
        temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
        if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + "Report2_page5_results.csv", 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report2_page5_results.csv")
        # SECTION 2 - By Org, Org dedup, AMR proportion
        if not AL.fn_savecsv(df_isoRep_blood, AC.CONST_PATH_RESULT + "Report2_AMR_proportion_table.csv", 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report2_AMR_proportion_table.csv")
        if not AL.fn_savecsv(df_isoRep_blood_byorg, AC.CONST_PATH_RESULT + "Report2_page6_counts_by_organism.csv", 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report2_page6_counts_by_organism.csv")
        if not AL.fn_savecsv(df_isoRep_blood_byorg_dedup, AC.CONST_PATH_RESULT + "Report2_page6_patients_under_this_surveillance_by_organism.csv", 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report2_page6_patients_under_this_surveillance_by_organism.csv")
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 3 - Summary
        if bishosp_ava:
            temp_list = [['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                         ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]],  
                         ['merged_data','Number_of_patients_with_blood_culture_positive_for_organism_under_this_survey', df_COHO_isoRep_blood_byorg["Number_of_patients_with_blood_culture_positive"].sum()], 
                         ['merged_data','Number_of_patients_with_community_origin_BSI', df_COHO_isoRep_blood_byorg["Community_origin"].sum()], 
                         ['merged_data','Number_of_patients_with_hospital_origin_BSI', df_COHO_isoRep_blood_byorg["Hospital_origin"].sum()], 
                         ['merged_data','Number_of_patients_with_unknown_origin_BSI', df_COHO_isoRep_blood_byorg["Unknown_origin"].sum()]]
            temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
            if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + "Report3_page12_results.csv", 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report3_page12_results.csv")
            # SECTION 3 - By Org, Org dedup, AMR proportion
            if not AL.fn_savecsv(df_COHO_isoRep_blood, AC.CONST_PATH_RESULT + "Report3_table.csv", 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report3_table.csv")
            if not AL.fn_savecsv(df_COHO_isoRep_blood_byorg, AC.CONST_PATH_RESULT + "Report3_page13_counts_by_origin.csv", 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report3_page13_counts_by_origin.csv")
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 4 - Summary
        if dict_progvar["n_blood_neg"] > 0:
            temp_list = [['merged_data','Minimum_date', dict_progvar["micro_date_min"]], 
                         ['merged_data','Maximum_date', dict_progvar["micro_date_max"]],  
                         ['merged_data','Number_of_blood_specimens_collected', dict_progvar["n_blood"]], 
                         ['merged_data','Number_of_patients_sampled_for_blood_culture', dict_progvar["n_blood_patients"]]]
            temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
            if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + "Report4_page24_results.csv", 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report4_page24_results.csv")
            # SECTION 4 By Org, Org and atb
            if not AL.fn_savecsv(df_isoRep_blood_incidence, AC.CONST_PATH_RESULT + "Report4_frequency_blood_samples.csv", 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report4_frequency_blood_samples.csv")
            if not AL.fn_savecsv(df_isoRep_blood_incidence_atb, AC.CONST_PATH_RESULT + "Report4_frequency_priority_pathogen.csv", 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report4_frequency_priority_pathogen.csv")
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 5 - Summary
        if bishosp_ava:
            if dict_progvar["n_blood_neg"] > 0:
                temp_df = df_hospmicro_blood[[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_CLEANADMDATE]].dropna().drop_duplicates()
                temp_df = temp_df.groupby([AC.CONST_NEWVARNAME_HN])[AC.CONST_NEWVARNAME_CLEANADMDATE].count().reset_index(name="count")
                temp_list = [['merged_data','Minimum_date', dict_progvar["micro_date_min"]], 
                             ['merged_data','Maximum_date', dict_progvar["micro_date_max"]],  
                             ['merged_data','Number_of_blood_specimens_collected', dict_progvar["n_blood"]], 
                             ['merged_data','Number_of_patients_sampled_for_blood_culture', dict_progvar["n_blood_patients"]],
                             ['merged_data','Number_of_patients_with_blood_culture_within_first_2_days_of_admission', dict_progvar["n_CO_blood_patients"]], 
                             ['merged_data','Number_of_patients_with_blood_culture_within_after_2_days_of_admission', dict_progvar["n_HO_blood_patients"]],
                             ['merged_data','Number_of_patients_with_unknown_origin',dict_progvar["n_blood_patients"] - len(df_hospmicro_blood[AC.CONST_NEWVARNAME_HN].unique()) ],
                             ['merged_data','Number_of_patients_had_more_than_one_admission', dict_progvar["n_2adm_firstbothCOHO_patients"]]]
                temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
                if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + "Report5_page27_results.csv", 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report5_page27_results.csv")
                # SECTION 5 By Org, Org and atb by CO/HO
                if not AL.fn_savecsv(df_COHO_isoRep_blood_incidence[df_COHO_isoRep_blood_incidence["Infection_origin"]==AC.CONST_EXPORT_COHO_CO_DATAVAL], AC.CONST_PATH_RESULT + "Report5_incidence_blood_samples_community_origin.csv", 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report5_incidence_blood_samples_community_origin.csv")
                if not AL.fn_savecsv(df_COHO_isoRep_blood_incidence_atb[df_COHO_isoRep_blood_incidence_atb["Infection_origin"]==AC.CONST_EXPORT_COHO_CO_DATAVAL], AC.CONST_PATH_RESULT + "Report5_incidence_blood_samples_community_origin_antibiotic.csv", 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report5_incidence_blood_samples_community_origin_antibiotic.csv")
                if not AL.fn_savecsv(df_COHO_isoRep_blood_incidence[df_COHO_isoRep_blood_incidence["Infection_origin"]==AC.CONST_EXPORT_COHO_HO_DATAVAL], AC.CONST_PATH_RESULT + "Report5_incidence_blood_samples_hospital_origin.csv", 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report5_incidence_blood_samples_hospital_origin.csv")
                if not AL.fn_savecsv(df_COHO_isoRep_blood_incidence_atb[df_COHO_isoRep_blood_incidence_atb["Infection_origin"]==AC.CONST_EXPORT_COHO_HO_DATAVAL], AC.CONST_PATH_RESULT + "Report5_incidence_blood_samples_hospital_origin_antibiotic.csv", 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report5_incidence_blood_samples_hospital_origin_antibiotic.csv")
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 6 - Summary
        if bishosp_ava:
            if dict_progvar["checkpoint_section6"] > 0:
                n_hosp_unique = len(df_hosp[AC.CONST_NEWVARNAME_HN_HOSP].unique())
                #n_hosp_died_unique = len(df_hosp[df_hosp[AC.CONST_NEWVARNAME_DISOUTCOME_HOSP]==AC.CONST_DIED_VALUE][AC.CONST_NEWVARNAME_HN_HOSP].unique())
                n_hosp_died_unique = len(df_hosp[df_hosp[AC.CONST_NEWVARNAME_DISOUTCOME_HOSP]==AC.CONST_DIED_VALUE])
                n_died_per = int(round((n_hosp_died_unique/n_hosp_unique)*100, 0))
                n_died_per_text = str(n_died_per) +"% (" + str(n_hosp_died_unique) + "/" + str(n_hosp_unique) + ")"
                temp_list = [['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                             ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]], 
                             ['merged_data','Number_of_blood_culture_positive_for_organism_under_this_survey', df_COHO_isoRep_blood_byorg["Number_of_patients_with_blood_culture_positive"].sum()], 
                             ['merged_data','Number_of_patients_with_community_origin_BSI', df_COHO_isoRep_blood_byorg["Community_origin"].sum()], 
                             ['merged_data','Number_of_patients_with_hospital_origin_BSI', df_COHO_isoRep_blood_byorg["Hospital_origin"].sum()],
                             ['hospital_admission_data','Minimum_date', dict_progvar["hosp_date_min"]], 
                             ['hospital_admission_data','Maximum_date', dict_progvar["hosp_date_max"]],
                             ['merged_data','Number_of_records', len(df_hosp) if bishosp_ava == True else 'NA'], 
                             ['merged_data','Number_of_patients_included', n_hosp_unique if bishosp_ava == True else 'NA'], 
                             ['merged_data','Number_of_deaths', n_hosp_died_unique if bishosp_ava == True else 'NA'], 
                             ['merged_data','Mortality', n_died_per_text if bishosp_ava == True else 'NA']
                             ]
                temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
                if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + "Report6_page32_results.csv", 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report6_page32_results.csv")
                # SECTION 6 By Org, Org and atb by CO/HO
                if not AL.fn_savecsv(df_COHO_isoRep_blood_mortality_byorg, AC.CONST_PATH_RESULT + "Report6_mortality_byorganism.csv", 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report6_mortality_byorganism.csv")
                if not AL.fn_savecsv(df_COHO_isoRep_blood_mortality, AC.CONST_PATH_RESULT + "Report6_mortality_table.csv", 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "Report6_mortality_table.csv")
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION ANNEX A - Summary
        if not AL.fn_savecsv(df_annexA, AC.CONST_PATH_RESULT + "AnnexA_page39_results.csv", 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "AnnexA_page39_results.csv")
        # A1
        if not AL.fn_savecsv(df_annexA1, AC.CONST_PATH_RESULT + "AnnexA_patients_with_positive_specimens.csv", 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "AnnexA_patients_with_positive_specimens.csv")
        # A2
        if bishosp_ava:
            if not AL.fn_savecsv(df_annexA2, AC.CONST_PATH_RESULT + "AnnexA_mortlity_table.csv", 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "AnnexA_mortlity_table.csv")
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # Data verification log
        #3.0.2
        """
        sspecdateformat = ""
        try:
            temp_df = df_micro[df_micro[AC.CONST_VARNAME_SPECDATERAW].isnull() == False]
            sspecdateformat = str(temp_df[AC.CONST_VARNAME_SPECDATERAW].max())
            sspecdateformat = sspecdateformat + ", " + str(temp_df[AC.CONST_VARNAME_SPECDATERAW].min())
            imid = int(len(temp_df)/2)
            sspecdateformat = sspecdateformat + ", " + str(temp_df[AC.CONST_VARNAME_SPECDATERAW][imid])
        except Exception as e: # work on python 3.x
            AL.printlog("Warning : Error get example of specimen date format " +  str(e),True,logger)  
            logger.exception(e)
        sadmdateformat = ""
        try:
            temp_df = df_hosp[df_hosp[AC.CONST_VARNAME_ADMISSIONDATE].isnull() == False]
            sadmdateformat = str(temp_df[AC.CONST_VARNAME_ADMISSIONDATE].max())
            sadmdateformat = sadmdateformat + ", " + str(temp_df[AC.CONST_VARNAME_ADMISSIONDATE].min())
            imid = int(len(temp_df)/2)
            sadmdateformat = sadmdateformat + ", " + str(temp_df[AC.CONST_VARNAME_ADMISSIONDATE][imid])
        except Exception as e: # work on python 3.x
            AL.printlog("Warning : Error get example of admission date format " +  str(e),True,logger)  
            logger.exception(e)
        """
        ispecnohn = 0
        ispecnotadm = 0
        imergeall = 0
        imergeblood = 0
        imergebsi = 0
        try:
            temp_df=fn_notindata(df_micro, df_datalog_mergedlist, AC.CONST_NEWVARNAME_MICROREC_ID, AC.CONST_NEWVARNAME_MICROREC_ID,True)
            ispecnohn = len(temp_df)
            temp_df=fn_notindata(df_datalog_mergedlist,df_hospmicro, AC.CONST_NEWVARNAME_MICROREC_ID, AC.CONST_NEWVARNAME_MICROREC_ID,True)
            ispecnotadm = len(temp_df)
            imergeall = len(df_hospmicro[AC.CONST_NEWVARNAME_MICROREC_ID].drop_duplicates())
            imergeblood = len(df_hospmicro_blood[AC.CONST_NEWVARNAME_MICROREC_ID].drop_duplicates())
            imergebsi=len(df_hospmicro_bsi[AC.CONST_NEWVARNAME_MICROREC_ID].drop_duplicates())
        except:
            pass
        AL.printlog("Complete export data for report section 1-6, Annex A: " + str(datetime.now()),False,logger)
        #Data logger
        temp_list = [['microbiology_data','Number_of_records', len(df_micro)], 
                     ['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                     ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]], 
                     ['microbiology_data','Number_of_missing_specimen_date',str(len(df_micro[df_micro[AC.CONST_VARNAME_SPECDATERAW].isnull()]))],
                     ['microbiology_data','Number_of_missing_specimen_type',str(len(df_micro[df_micro[AC.CONST_VARNAME_SPECTYPE].isnull()]))],
                     ['microbiology_data','Number_of_missing_culture_result',str(len(df_micro[df_micro[AC.CONST_VARNAME_ORG].isnull()]))],
                     ['hospital_admission_data','Number_of_records', len(df_hosp) if bishosp_ava else "NA"], 
                     ['hospital_admission_data','Minimum_date', dict_progvar["hosp_date_min"] if bishosp_ava else "NA"], 
                     ['hospital_admission_data','Maximum_date', dict_progvar["hosp_date_max"] if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_admission_date',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_ADMISSIONDATE].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_discharge_date',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_DISCHARGEDATE].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_outcome_result',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_DISCHARGESTATUS].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_age',str(len(df_hosp[df_hosp[AC.CONST_NEWVARNAME_AGEYEAR].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_gender',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_GENDER].isnull()])) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_unmatchhn',str(ispecnohn) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_unmatchperiod',str(ispecnotadm) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_matchall',str(imergeall) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_matchblood',str(imergeblood) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_matchbsi',str( imergebsi) if bishosp_ava else "NA"],
                     ['microbiology_data','Hospital_name',dict_amasstodataval["hospital_name"]],
                     ['microbiology_data','Country',dict_amasstodataval["country"]]
                     ]
        """
        temp_list = [['microbiology_data','Hospital_name',dict_amasstodataval["hospital_name"]],
                    ['microbiology_data','Country',dict_amasstodataval["country"]],
                    ['microbiology_data','Minimum_date',dict_progvar["micro_date_min"]],
                    ['microbiology_data','Maximum_date',dict_progvar["micro_date_max"]],
                    ['microbiology_data','Number_of_missing_specimen_date',str(len(df_micro[df_micro[AC.CONST_VARNAME_SPECDATERAW].isnull()]))],
                    ['microbiology_data','Number_of_missing_specimen_type',str(len(df_micro[df_micro[AC.CONST_VARNAME_SPECTYPE].isnull()]))],
                    ['microbiology_data','Number_of_missing_culture_result',str(len(df_micro[df_micro[AC.CONST_VARNAME_ORG].isnull()]))],
                    ['microbiology_data','format_of_specimen_date',"[" + df_micro.loc[1, AC.CONST_NEWVARNAME_CLEANSPECDATE].strftime("%d-%b-%Y") + "] [" + df_micro.loc[2, AC.CONST_NEWVARNAME_CLEANSPECDATE].strftime("%d-%b-%Y") + "]"],
                    ['hospital_admission_data','Number_of_missing_admission_date',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_ADMISSIONDATE].isnull()])) if bishosp_ava else "NA"],
                    ['hospital_admission_data','Number_of_missing_discharge_type',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_DISCHARGEDATE].isnull()])) if bishosp_ava else "NA"],
                    ['hospital_admission_data','Number_of_missing_outcome_result',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_DISCHARGESTATUS].isnull()])) if bishosp_ava else "NA"],
                    ['hospital_admission_data','format_of_admission_date',"[" + df_hosp.loc[1, AC.CONST_NEWVARNAME_CLEANADMDATE].strftime("%d-%b-%Y") + "] [" + df_hosp.loc[2, AC.CONST_NEWVARNAME_CLEANADMDATE].strftime("%d-%b-%Y") + "]" if bishosp_ava else "[NA] [NA]"],
                    ['hospital_admission_data','format_of_discharge_date',"[" + df_hosp.loc[1, AC.CONST_NEWVARNAME_CLEANDISDATE].strftime("%d-%b-%Y") + "] [" + df_hosp.loc[2, AC.CONST_NEWVARNAME_CLEANDISDATE].strftime("%d-%b-%Y") + "]" if bishosp_ava else "[NA] [NA]"],
                    ['merged_data','Number_of_missing_specimen_date',str(len(df_hospmicro_bsi[df_hospmicro_bsi[AC.CONST_NEWVARNAME_CLEANSPECDATE].isnull()]))],
                    ['merged_data','Number_of_missing_admission_date',str(len(df_hospmicro_bsi[df_hospmicro_bsi[AC.CONST_NEWVARNAME_CLEANADMDATE].isnull()]))],
                    ['merged_data','Number_of_missing_discharge_type',str(len(df_hospmicro_bsi[df_hospmicro_bsi[AC.CONST_NEWVARNAME_CLEANDISDATE].isnull()])) if bishosp_ava else "NA"],
                    ['merged_data','Number_of_missing_age',str(len(df_hospmicro_bsi[df_hospmicro_bsi[AC.CONST_NEWVARNAME_AGEYEAR].isnull()])) if bishosp_ava else "NA"],
                    ['merged_data','Number_of_missing_gender',str(len(df_hospmicro_bsi[df_hospmicro_bsi[AC.CONST_VARNAME_GENDER].isnull()])) if bishosp_ava else "NA"],
                    ['merged_data','Number_of_missing_infection_origin_data',str(len(df_hospmicro_bsi[df_hospmicro_bsi[AC.CONST_NEWVARNAME_COHO_FROMHOS].isnull()])) if bishosp_ava else "NA"],
                    ['example_specimen_date',sspecdateformat],
                    ['example_admission_date',sadmdateformat if bishosp_ava else "NA"]
                    ]
        """
        temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
        if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + "logfile_results.csv",2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "logfile_results.csv")
        temp_df = df_micro.groupby([AC.CONST_VARNAME_ORG])[AC.CONST_VARNAME_ORG].count().reset_index(name='Frequency')
        temp_df.rename(columns={AC.CONST_VARNAME_ORG: 'Organism'}, inplace=True)
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_organism.xlsx", logger):
            print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_organism.xlsx")
        temp_df = df_micro.groupby([AC.CONST_VARNAME_SPECTYPE])[AC.CONST_VARNAME_SPECTYPE].count().reset_index(name='Frequency')
        temp_df.rename(columns={AC.CONST_VARNAME_SPECTYPE: 'Specimen'}, inplace=True)
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_specimen.xlsx", logger):
            print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_specimen.xlsx")
        temp_df = check_date_format(df_micro, AC.CONST_VARNAME_SPECDATERAW, logger)
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_formatspecdate.xlsx", logger):
            print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_formatspecdate.xlsx")
        temp_df = check_hn_format(df_micro, AC.CONST_NEWVARNAME_HN, logger)
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_microhn.xlsx", logger):
            print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_microhn.xlsx")
        if bishosp_ava:
            temp_df = df_hosp.groupby([AC.CONST_VARNAME_GENDER])[AC.CONST_VARNAME_GENDER].count().reset_index(name='Frequency')
            temp_df.rename(columns={AC.CONST_VARNAME_GENDER: 'Gender'}, inplace=True)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_gender.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_gender.xlsx")
            temp_df = df_hosp.groupby([AC.CONST_VARNAME_DISCHARGESTATUS])[AC.CONST_VARNAME_DISCHARGESTATUS].count().reset_index(name='Frequency')
            temp_df.rename(columns={AC.CONST_VARNAME_DISCHARGESTATUS: 'Discharge status'}, inplace=True)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_discharge.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_discharge.xlsx")
            temp_df = df_hosp.groupby([AC.CONST_NEWVARNAME_AGEYEAR])[AC.CONST_NEWVARNAME_AGEYEAR].count().reset_index(name='Frequency')
            temp_df.rename(columns={AC.CONST_NEWVARNAME_AGEYEAR: 'Age'}, inplace=True)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_age.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_age.xlsx")
            temp_df = check_date_format(df_hosp, AC.CONST_VARNAME_ADMISSIONDATE, logger)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_formatadmdate.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_formatadmdate.xlsx")
            temp_df = check_date_format(df_hosp, AC.CONST_VARNAME_DISCHARGEDATE, logger)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_formatdisdate.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_formatdisdate.xlsx")
            temp_df = check_hn_format(df_hosp, AC.CONST_NEWVARNAME_HN_HOSP, logger)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_hosphn.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_hosphn.xlsx")
        AL.printlog("Complete Export data for data verifcation log (END analysis): " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail export data section 1-6, Appendix A: " +  str(e),True,logger)  
        logger.exception(e)
    AL.printlog("End AMR analysis: " + str(datetime.now()),False,logger)
    
    ANNEX_B.generate_annex_b(df_dict_micro, df_micro_annexb,logger,AC.CONST_ANNEXB_USING_MAPPEDDATA,False)
    #AMR_REPORT.generate_amr_report(df_dict_micro, logger)
    AMR_REPORT_NEW.generate_amr_report(df_dict_micro,dict_orgcatwithatb,dict_orgwithatb_mortality,dict_orgwithatb_incidence,bishosp_ava,logger)
    SUP_REPORT.generate_supplementary_report(df_dict_micro,logger,AC.CONST_ANNEXB_USING_MAPPEDDATA)
# ------------------------------------------------------------------------------------------------------------------------------------------------------      
#Main loop
mainloop()