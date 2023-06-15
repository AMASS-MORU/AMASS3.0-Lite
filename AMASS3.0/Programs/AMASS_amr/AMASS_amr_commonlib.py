#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** COMMON FUNCTION LIB                                                                             ***#
#***-------------------------------------------------------------------------------------------------***#
# @author: PRAPASS WANNAPINIJ
# Created on: 09 MAR 2023 
import pandas as pd #for creating and manipulating dataframe
import os
import csv
from xlsx2csv import Xlsx2csv
import logging #for creating error_log
# function for init log file
def initlogger(sprogname,slogname) :
    # Create a logging instance
    file_logger = logging.getLogger(sprogname)
    file_logger.setLevel(logging.INFO)
    # Assign a file-handler to that instance
    fh = logging.FileHandler(slogname)
    fh.setLevel(logging.INFO)
    # Format your logs (optional)
    fh.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
    # Add the handler to your logging instance
    file_logger.addHandler(fh)
    return file_logger
# Print to console and log to log file   
def printlog(smessage,biserror,logger) :
    try:
        print(smessage)
        if biserror:
            logger.error(smessage)
        else:
            logger.info(smessage)
        return True
    except:
        return False
# Check that is there either CSV or XLSX with the specific file name or not
def checkxlsorcsv(spath,sfilename) :
    bfound = False
    try:
        if os.path.exists(spath + sfilename + ".csv") :
            bfound = True
        else:
            if os.path.exists(spath + sfilename + ".xlsx") :
                bfound = True
    except:
        bfound = False
    return bfound
# Read csv or xlsx file, csv first, if no file convert xlsx to csv beflor load
def readxlsxorcsv(spath,sfilename) :
    df = pd.DataFrame()
    try:
        df = pd.read_csv(spath + sfilename + ".csv").fillna("")
    except:
        try:
            df= pd.read_csv(spath + sfilename + ".csv", encoding="windows-1252").fillna("")
        except:
            df= pd.read_excel(spath + sfilename + ".xlsx").fillna("")
            """
            stmpfilename = spath + sfilename + "_temp.csv"
            if os.path.exists(stmpfilename):
                os.remove(stmpfilename)
            if os.path.exists(spath + sfilename + ".xlsx") :
                #Xlsx2csv(spath + sfilename + ".xlsx", outputencoding="utf-8").convert(stmpfilename,sheetid=1)
                Xlsx2csv(spath + sfilename + ".xlsx").convert(stmpfilename,sheetid=1)
                df = pd.read_csv(stmpfilename)
            if os.path.exists(stmpfilename):
               os.remove(stmpfilename)
            """
    return df
# Read csv or xlsx file and specify header in columnheader list, csv first, if no file convert xlsx to csv beflor load
def readxlsorcsv_noheader(spath,sfilename,columnheader) :
    try:
        df = pd.read_csv(spath + sfilename + ".csv",header=None,names=columnheader).fillna("")
    except:
        try:
            df= pd.read_csv(spath + sfilename + ".csv", encoding="windows-1252",header=None,names=columnheader).fillna("")
        except:
            df= pd.read_excel(spath + sfilename + ".xlsx",header=None,names=columnheader).fillna("")
            """
            stmpfilename = spath + sfilename + "_temp.csv"
            if os.path.exists(stmpfilename):
                os.remove(stmpfilename)
            #Xlsx2csv(spath + sfilename + ".xlsx", outputencoding="utf-8").convert(stmpfilename,sheetid=1)    
            Xlsx2csv(spath + sfilename + ".xlsx").convert(stmpfilename,sheetid=1)
            df = pd.read_csv(stmpfilename,header=None,names=columnheader)
            if os.path.exists(stmpfilename):
                os.remove(stmpfilename)
            """
    return df
# Save to csv
def fn_savecsv(df,fname,iquotemode,logger) :
    try:
        if iquotemode == 1:
            df.to_csv(fname,index=False, quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
        elif iquotemode == 2:
            df.to_csv(fname,index=False, quotechar='"', quoting=csv.QUOTE_ALL)  
        else:
            df.to_csv(fname,index=False)
        #logger.info("save csv file : " + fname)
        return True
    except Exception as e: # work on python 3.x
        printlog("Failed to save csv file : " + fname + " : "+ str(e),True,logger)
        return False
# Save to xlsx
def fn_savexlsx(df,fname,logger) :
    try:
        df.to_excel(fname,index=False)
        #logger.info("save xlsx file : " + fname)
        return True
    except Exception as e: # work on python 3.x
        printlog("Failed to save xlsx file : " + fname + " : "+ str(e),True,logger)
        return False
# Remove specific columns in dataframe
def fn_removecol(df,list_col) :
    list_colexist = []
    for scol in df.columns:
        if scol not in list_col:
            list_colexist.append(scol)    
    if len(list_colexist) > 0:
        return df[list_colexist]
    else:
        return df    
# Remove unused columns in dataframe 
def fn_keeponlycol(df,list_col) :
    list_colexist = []
    for scol in list_col:
        if scol in df.columns:
            list_colexist.append(scol)
    if len(list_colexist) > 0:
        return df[list_colexist]
    else:
        return df
# Change field type to category to save mem space
def fn_df_tocategory_datatype(df,list_col,logger) :
    curcol = ""
    for scol in list_col:
        curcol = scol
        if scol in df.columns:
            try:
                df[scol] = df[scol].astype("category")
                printlog("Note : convert field type to category for column : " + curcol,False,logger)  
            except Exception as e: # work on python 3.x
                printlog("Warning : convert field type to category type : " + curcol+ " : "+ str(e),False,logger)         
    return df
# Trim space and unreadable charecters
def fn_df_strstrips(df,list_col,logger) :
    curcol = ""
    for scol in list_col:
        curcol = scol
        if scol in df.columns:
            try:
                df[scol] = df[scol].astype("string").str.strip()
                printlog("Note : trimmed column : " + curcol,False,logger)  
            except Exception as e: # work on python 3.x
                printlog("Warning : unable to trim data in column : " + curcol+ " : "+ str(e),False,logger)         
    return df