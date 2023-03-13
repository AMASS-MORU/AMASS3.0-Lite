#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate Data verification logfile reports systematically.

# Created on 20th April 2022
import logging #for creating logfile
import pandas as pd #for creating and manipulating dataframe
from datetime import date #for generating today date
from reportlab.lib.pagesizes import A4 #for setting PDF size
from reportlab.pdfgen import canvas #for creating PDF page
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.platypus import * #for plotting graph and tables
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
from AMASS_logfile_function_version_2 import * #for importing logfile functions
from reportlab.lib.colors import * #for importing color palette
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables


logger = logging.getLogger('AMASS_logfile_version_2.py')
logger.setLevel(logging.INFO)
fh = logging.FileHandler("./error_logfile_amass.txt")
fh.setLevel(logging.ERROR)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)

add_blankline = "<br/>"
today = date.today()
path = "./"
dict_i = path + "dictionary_for_microbiology_data.xlsx"
dict_hosp_i = path + "dictionary_for_hospital_admission_data.xlsx"
over_i = path + "ResultData/logfile_results.csv"
org_i = path + "ResultData/logfile_organism.xlsx"
spc_i = path + "ResultData/logfile_specimen.xlsx"
ast_i = path + "ResultData/logfile_ast.xlsx"
gen_i = path + "ResultData/logfile_gender.xlsx"
age_i = path + "ResultData/logfile_age.xlsx"
dis_i = path + "ResultData/logfile_discharge.xlsx"
var_mi_i = path + "Variables/variables_micro.xlsx"
var_ho_i = path + "Variables/variables_hosp.xlsx"


try:
    over_raw = pd.read_csv(over_i)
except:
    over_raw = pd.DataFrame(["microbiology_data","microbiology_data","microbiology_data","microbiology_data", 
                             "hospital_admission_data","hospital_admission_data","hospital_admission_data","hospital_admission_data","hospital_admission_data", 
                             "merged_data","merged_data","merged_data","merged_data","merged_data","merged_data"],index=range(15), columns=["Type_of_data_file"])
    over_raw["Parameters"] =["Number_of_missing_specimen_date","Number_of_missing_specimen_type","Number_of_missing_culture_result","format_of_specimen_date",
                             "Number_of_missing_admission_date","Number_of_missing_discharge_type","Number_of_missing_outcome_result","format_of_admission_date","format_of_discharge_date", 
                             "Number_of_missing_specimen_date","Number_of_missing_admission_date","Number_of_missing_discharge_type","Number_of_missing_age","Number_of_missing_gender","Number_of_missing_infection_origin_data"]
    over_raw["Values"] = "NA"

if checkpoint(dict_i):
    try:
        try:
            dict_raw = pd.read_excel(dict_i).iloc[:,:4]
        except:
            try:
                dict_raw = pd.read_csv(path + "dictionary_for_microbiology_data.csv").iloc[:,:4]
            except:
                dict_raw = pd.read_csv(path + "dictionary_for_microbiology_data.csv",encoding="windows-1252").iloc[:,:4]
        dict_raw.columns = ["amass_name","user_name","requirements","explanations"]
        file_format = retrieve_uservalue(dict_raw, "file_format")
        lst_opt_2 = ["1", 
                    "hospital_number", retrieve_uservalue(dict_raw, "hospital_number"), 
                    "specimen_collection_date", retrieve_uservalue(dict_raw, "specimen_collection_date"), 
                    "specimen_type", retrieve_uservalue(dict_raw, "specimen_type"), 
                    "organism", retrieve_uservalue(dict_raw, "organism"), 
                    "specimen_number", retrieve_uservalue(dict_raw, "specimen_number"), 
                    "antibiotic", retrieve_uservalue(dict_raw, "antibiotic"), 
                    "ast_result", retrieve_uservalue(dict_raw, "ast_result")]
        lst_opt_2 = [x for x in lst_opt_2 if pd.isna(x)==False] #selecting only variables which are not NaN
        #Retrieving all antibiotic list from dictionary_for_microbiology_data.xlsx
        idx_micro_drug = dict_raw.iloc[:,0].tolist().index("Variable names used for \"antibiotics\" in AMASS") #index of part2 header
        idx_micro_spc = dict_raw.iloc[:,0].tolist().index("Data values of variable used for \"specimen_type\" in AMASS") #index of part3 header
        idx_micro_org_core =  dict_raw.iloc[:,0].tolist().index("Data values of variable \"organism\" in AMASS, which are mainly used for the main report")
        idx_micro_opt =  dict_raw.iloc[:,0].tolist().index("Optional: Data values used for the cover of the report generated by the AMASS")
        idx_micro_org_opt =  dict_raw.iloc[:,0].tolist().index("Optional: Data values of variable \"organism\" in AMASS, which are mainly used for the annex")
        idx_micro_opt_2 =  dict_raw.iloc[:,0].tolist().index("Optional: Variables name used in AMASS; this is needed only when data is in long format")

        dict_drug = dict_raw.copy().iloc[idx_micro_drug+1:idx_micro_spc,:2].reset_index().drop(columns=['index']).fillna("").rename(columns={"amass_name":"amass_drug","user_name":"user_drug"})
        dict_spc = dict_raw.copy().iloc[idx_micro_spc+1:idx_micro_org_core,:2].reset_index().drop(columns=['index']).fillna("")
        dict_org_core = dict_raw.copy().iloc[idx_micro_org_core+1:idx_micro_opt,:2].reset_index().drop(columns=['index']).fillna("")
        dict_org_core["tag"] = "A"
        dict_org_opt = dict_raw.copy().iloc[idx_micro_org_opt+1:idx_micro_opt_2,:2].reset_index().drop(columns=['index']).fillna("")
        dict_org_opt["tag"] = "B"
        dict_org = pd.concat([dict_org_core,dict_org_opt],axis=0)
    except Exception as e:
        logger.exception(e)
        pass

    marked_idx_org = []
    try: #TS1
        try:
            org_raw = pd.read_excel(org_i)
        except:
            org_raw = pd.read_csv(org_i,encoding='windows-1252')
        org_merge = pd.merge(dict_org,org_raw,how="outer",left_on="user_name",right_on="Organism")
        org_merge["Frequency"] = org_merge["Frequency"].fillna(0).astype(int).astype(str)
        org_merge["amass_name"] = org_merge["amass_name"].fillna("zzzz")
        org_merge = org_merge.fillna("zzzz").replace("","zzzz").sort_values(["amass_name"],ascending=True) #all NaN and unknown values >>>> "zzzz" for ordering organism name
        org_A = org_merge.loc[(org_merge["tag"]=="A")&(org_merge["amass_name"]!="zzzz")].sort_values(["amass_name"],ascending=True).replace("zzzz","").reset_index().drop(columns=["index","tag"])
        for idx in org_A.index:
            if org_A.loc[idx,"user_name"] != "" and org_A.loc[idx,"Organism"] == "":
                org_A.at[idx,"Organism"] = org_A.loc[idx,"user_name"]
            else:
                pass
            org_A.at[idx,"Organism"] = prepare_unicode(org_A.loc[idx,"Organism"])
        org_A = org_A.drop(columns=["user_name"]).rename(columns={"amass_name":"Data values of variable \"organism\"\nin AMASS, which are mainly used for\nthe main report","Organism":"Data values of variable\nrecorded for \"organism\" in your\nmicrobiology data file", "Frequency":"Number of\nobservations"})
        marked_org_A = marked_idx(org_A)
        org_A_col = [list(org_A.columns)]

        org_B = org_merge.loc[(org_merge["tag"]=="B")|(org_merge["amass_name"]=="zzzz"),:].sort_values(["amass_name"],ascending=True).replace("zzzz","").reset_index().drop(columns=["index","tag"])
        for idx in org_B.index:
            if org_B.loc[idx,"user_name"] != "" and org_B.loc[idx,"Organism"] == "":
                org_B.at[idx,"Organism"] = org_B.loc[idx,"user_name"]
            else:
                pass
            org_B.at[idx,"Organism"] = prepare_unicode(org_B.loc[idx,"Organism"])
        org_B = org_B.drop(columns=["user_name"]).rename(columns={"amass_name":"Optional: Data values of variable\n\"organism\" in AMASS, which are \nmainly used for the annex", "Organism":"Data values of the variable\nrecorded for \"organism\" in your\nmicrobiology data file", "Frequency":"Number of\nobservations"})
        marked_org_B = marked_idx(org_B)
        org_B_col = [list(org_B.columns)]
    except Exception as e:
        logger.exception(e)
        pass

    marked_idx_spc = []
    try: #TS2
        try:
            spc_raw = pd.read_excel(spc_i)
        except:
            spc_raw = pd.read_csv(spc_i,encoding='windows-1252')
        spc_merge = pd.merge(dict_spc,spc_raw,how="outer",left_on="user_name",right_on="Specimen")
        spc_merge["Frequency"] = spc_merge["Frequency"].fillna(0).astype(int).astype(str)
        spc_merge = spc_merge.fillna("zzzz")
        spc_merge = spc_merge.sort_values(["amass_name"],ascending=True).reset_index().drop(columns=["index"])
        spc_merge = spc_merge.replace(regex=["zzzz"],value="")
        for idx in spc_merge.index:
            if spc_merge.loc[idx,"user_name"] != "" and spc_merge.loc[idx,"Specimen"] == "":
                spc_merge.at[idx,"Specimen"] = spc_merge.loc[idx,"user_name"]
            else:
                pass
            spc_merge.at[idx,"Specimen"] = prepare_unicode(spc_merge.loc[idx,"Specimen"])
        spc_merge = spc_merge.drop(columns=["user_name"]).fillna("").rename(columns={"amass_name":"Data values of variable used\nfor \"specimen_type\" in AMASS",
                                                "Specimen":"Data values of variable\nrecorded for \"specimen_type\" in your\nmicrobiology data file", 
                                                "Frequency":"Number of\nobservations"})
        marked_spc = marked_idx(spc_merge)
        spc_col = [list(spc_merge.columns)]
    except Exception as e:
        logger.exception(e)
        pass

    try: #TS3
        try:
            ast_raw = pd.read_excel(ast_i)
        except:
            ast_raw = pd.read_csv(ast_i,encoding='windows-1252')
        ast_raw.columns = ["user_name", "frequency_raw"]
        ast_merge = pd.merge(dict_drug,ast_raw,how="outer",left_on="amass_drug",right_on="user_name")
        ast_merge[["user_drug","user_name"]] = ast_merge[["user_drug","user_name"]].fillna("")
        for idx in ast_merge.index:
            if ast_merge.loc[idx,"amass_drug"] == ast_merge.loc[idx,"user_name"]:
                pass
            else:
                if ast_merge.loc[idx,"user_drug"] != "" and ast_merge.loc[idx,"user_name"] == "":
                    pass
                else:
                    ast_merge.at[idx,"user_drug"] = ast_merge.loc[idx,"user_name"]
        ast_merge["frequency_raw"] = ast_merge["frequency_raw"].fillna(0).astype(int).astype(str)
        ast_merge["amass_drug"] = ast_merge["amass_drug"].fillna("zzzz")
        ast_merge = ast_merge.sort_values(["amass_drug"],ascending=True).reset_index().drop(columns=["index"])
        ast_merge["amass_drug"] = ast_merge["amass_drug"].replace(regex=["zzzz"],value="")
        ast_merge = ast_merge.loc[~ast_merge["user_drug"].isin(lst_opt_2),:] #excluding variables out
        for idx in ast_merge.index:
            ast_merge.at[idx,"user_drug"] = prepare_unicode(ast_merge.loc[idx,"user_drug"])
        ast_merge = ast_merge.reset_index().drop(columns=["index","user_name"]).rename(columns={"amass_drug":"Variable names used for\n\"antibiotics\" described in AMASS",
                                                                        "user_drug":"Variable names recorded for\n\"antibiotics\" in your\nmicrobiology data file",
                                                                        "frequency_raw":"Number of observations\ncontaining S, I, or R\nfor each antibiotic"})
        ast_col = [list(ast_merge.columns)]
        marked_ast = marked_idx(ast_merge)
    except Exception as e:
        logger.exception(e)
        pass

##dictionary_for_hospital_admission_data
if checkpoint(dict_hosp_i): 
    try:
        try:
            dict_hosp_raw = pd.read_excel(dict_hosp_i).iloc[:,:4]
        except:
            try:
                dict_hosp_raw = pd.read_csv(path + "dictionary_for_hospital_admission_data.csv").iloc[:,:4]
            except:
                dict_hosp_raw = pd.read_csv(path + "dictionary_for_hospital_admission_data.csv",encoding="windows-1252").iloc[:,:4]
        dict_hosp_raw.columns = ["amass_name","user_name","requirements","explanations"]
        #Retrieving column names of hospital_admission_data.xlsx 
        male = dict_hosp_raw.loc[dict_hosp_raw["amass_name"]=="male",["amass_name","user_name"]]
        female = dict_hosp_raw.loc[dict_hosp_raw["amass_name"]=="female",["amass_name","user_name"]]
        dict_gen = pd.concat([male,female],axis=0)
        died = dict_hosp_raw.loc[dict_hosp_raw["amass_name"]=="died",["amass_name","user_name"]]
    except Exception as e:
        logger.exception(e)
        pass

    try: #TS4
        try:
            gen_raw = pd.read_excel(gen_i)
        except:
            gen_raw = pd.read_csv(gen_i,encoding='windows-1252')
        gen_merge = pd.merge(dict_gen,gen_raw,how="outer",left_on="user_name",right_on="Gender")
        gen_merge["Frequency"] = gen_merge["Frequency"].fillna(0).astype(int).astype(str)
        for idx in gen_merge.index:
            if gen_merge.loc[idx,"user_name"] != "" and gen_merge.loc[idx,"Gender"] == "":
                gen_merge.at[idx,"Gender"] = gen_merge.loc[idx,"user_name"]
            else:
                pass
            gen_merge.at[idx,"Gender"] = prepare_unicode(gen_merge.loc[idx,"Gender"])
        gen_merge = gen_merge.drop(columns=["user_name"]).fillna("").rename(columns={"amass_name":"Data values of variable\n used for \"gender\" described\nin AMASS",
                                                        "Gender":"Data values of variable\n recorded for \"gender\" in your\nhospital admission data file", 
                                                        "Frequency":"Number of\nobservations"})
        gen_col = [list(gen_merge.columns)]
        gen_1 = gen_merge.values.tolist()
        gen_1 = gen_col + gen_1
        marked_gen = marked_idx(gen_merge)
    except Exception as e:
        logger.exception(e)
        pass

    try: #TS5
        try:
            age_raw = pd.read_excel(age_i)
            age_raw["Age"] = age_raw["Age"].fillna("NA")
        except:
            age_raw = pd.read_csv(age_i,encoding='windows-1252')
        age_raw["Age_cat"] = ""
        for idx in age_raw.index:
            if age_raw.loc[idx,"Age"] == "NA":
                age_raw.at[idx,"Age_cat"] = "Not available"
            elif int(age_raw.loc[idx,"Age"]) < 1:
                age_raw.at[idx,"Age_cat"] = "Less than 1 year"
            elif int(age_raw.loc[idx,"Age"]) >= 1 and int(age_raw.loc[idx,"Age"]) <= 4:
                age_raw.at[idx,"Age_cat"] = "1 to 4 years"
            elif int(age_raw.loc[idx,"Age"]) >= 5 and int(age_raw.loc[idx,"Age"]) <= 14:
                age_raw.at[idx,"Age_cat"] = "5 to 14 years"
            elif int(age_raw.loc[idx,"Age"]) >= 15 and int(age_raw.loc[idx,"Age"]) <= 24:
                age_raw.at[idx,"Age_cat"] = "15 to 24 years"
            elif int(age_raw.loc[idx,"Age"]) >= 25 and int(age_raw.loc[idx,"Age"]) <= 34:
                age_raw.at[idx,"Age_cat"] = "25 to 34 years"
            elif int(age_raw.loc[idx,"Age"]) >= 35 and int(age_raw.loc[idx,"Age"]) <= 44:
                age_raw.at[idx,"Age_cat"] = "35 to 44 years"
            elif int(age_raw.loc[idx,"Age"]) >= 45 and int(age_raw.loc[idx,"Age"]) <= 54:
                age_raw.at[idx,"Age_cat"] = "45 to 54 years"
            elif int(age_raw.loc[idx,"Age"]) >= 55 and int(age_raw.loc[idx,"Age"]) <= 64:
                age_raw.at[idx,"Age_cat"] = "55 to 64 years"
            elif int(age_raw.loc[idx,"Age"]) >= 65 and int(age_raw.loc[idx,"Age"]) <= 80:
                age_raw.at[idx,"Age_cat"] = "65 to 80 years"
            else:
                age_raw.at[idx,"Age_cat"] = "More than 80 years"
        age_raw["Frequency"] = age_raw["Frequency"].fillna(0).astype(int).astype(str)
        age_merge = age_raw.copy().loc[:,["Age_cat","Age","Frequency"]].fillna("").rename(columns={"Age_cat":"Data values of variable\n used for \"age\" described\nin AMASS",
                                                                                            "Age":"Data values of variable\nrecorded for \"age\" in your\nhospital admission data file", 
                                                                                            "Frequency":"Number of\nobservations"})
        age_col = [list(age_merge.columns)]
        marked_age = marked_idx(age_merge)
    except Exception as e:
        logger.exception(e)
        pass

    try: #TS6
        try:
            dis_raw = pd.read_excel(dis_i)
        except:
            dis_raw = pd.read_csv(dis_i,encoding='windows-1252')
        dis_merge = pd.merge(died,dis_raw,how="outer",left_on="user_name",right_on="Discharge status")
        dis_merge["Frequency"] = dis_merge["Frequency"].fillna(0).astype(int).astype(str)
        for idx in dis_merge.index:
            if dis_merge.loc[idx,"user_name"] != "" and dis_merge.loc[idx,"Discharge status"] == "":
                dis_merge.at[idx,"Discharge status"] = dis_merge.loc[idx,"user_name"]
            else:
                pass
            dis_merge.at[idx,"Discharge status"] = prepare_unicode(dis_merge.loc[idx,"Discharge status"])
        dis_merge = dis_merge.drop(columns=["user_name"]).fillna("").rename(columns={"amass_name":"Data values of variable\nname used for \"discharge status\"\ndescribed in AMASS",
                                                        "Discharge status":"Data values of variable\nname recorded for \"discharge status\"\nin your hospital admission data file", 
                                                        "Frequency":"Number of\nobservations"})
        dis_col = [list(dis_merge.columns)]
        marked_dis = marked_idx(dis_merge)
    except Exception as e:
        logger.exception(e)
        pass


if checkpoint(var_mi_i): #TS7
    try:
        try:
            var_mi_raw = pd.read_excel(var_mi_i).rename(columns={"variables_micro":"Variable names used in your microbiology data file"})
        except:
            var_mi_i = path+"Variables/variables_micro.csv"
            var_mi_raw = pd.read_csv(var_mi_i,encoding='windows-1252').rename(columns={"variables_micro":"Variable names used in your microbiology data file"})
        var_mi_col = [list(var_mi_raw.columns)]
        marked_var_mi = marked_idx(var_mi_raw)
    except Exception as e:
        logger.exception(e) # Will send the errors to the file
        pass

if checkpoint(var_ho_i): #TS8
    try: #TS8
        try:
            var_ho_raw = pd.read_excel(var_ho_i).rename(columns={"variables_hosp":"Variable names used in your hospital admission data file"})
        except:
            var_ho_i = path+"Variables/variables_hosp.csv"
            var_ho_raw = pd.read_csv(var_ho_i,encoding='windows-1252').rename(columns={"variables_hosp":"Variable names used in your hospital admission data file"})
        var_ho_col = [list(var_ho_raw.columns)]
        marked_var_ho = marked_idx(var_ho_raw)
    except Exception as e:
        logger.exception(e)
        pass

def cover(over_raw, today=today.strftime("%d %b %Y")):
    ##paragraph variable
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    add_blankline = "<br/>"
    hospital_name = assign_na_toinfo(retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Hospital_name","Parameters"),coverpage=True)
    country_name  = assign_na_toinfo(retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Country","Parameters"),coverpage=True)
    spc_date_start_cov = assign_na_toinfo(retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Minimum_date","Parameters"),coverpage=True)
    spc_date_end_cov   = assign_na_toinfo(retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Maximum_date","Parameters"),coverpage=True)    ##content
    cover_1_1 = "<b>Hospital name:</b>  " + bold_blue_op + hospital_name + bold_blue_ed
    cover_1_2 = "<b>Country name:</b>  " + bold_blue_op + country_name + bold_blue_ed
    cover_1_3 = "<b>Data from:</b>"
    cover_1_4 = bold_blue_op + str(spc_date_start_cov) + " to " + str(spc_date_end_cov) + bold_blue_ed
    cover_1 = [cover_1_1,cover_1_2,add_blankline+cover_1_3, cover_1_4]
    cover_2_1 = "This report is for users to review variable names and data values used by microbiology_data file and hospital_admission_data file saved within the same folder as the application file (AMASS.bat). This report can be used to assist users while completing the data dictionaries for both data files."
    cover_2_2 = "<b>Generated on:</b>  " + bold_blue_op + today + bold_blue_ed
    cover_2 = [cover_2_1,cover_2_2]
    ##reportlab
    c.setFillColor('#FCBB42')
    c.rect(0,590,800,20, fill=True, stroke=False)
    c.setFillColor(grey)
    c.rect(0,420,800,150, fill=True, stroke=False)
    report_title(c,'Data verification log file',0.7*inch, 485,'white',font_size=28)
    report_context(c,cover_1, 0.7*inch, 3.0*inch, 460, 180, font_size=18,line_space=26)
    report_context(c,cover_2, 0.7*inch, 0.5*inch, 460, 120, font_size=11)
    c.showPage()

def ts0(over_raw):
    log_1_1 = ""
    log_1_2 = "Please review the following information carefully before interpreting the AMR surveillance report generated by the AMASS application."
    log_1_3 = "Missing data− Microbiology data set (file name: microbiology_data)"
    log_1_4 = "The number of observations with missing specimen date: "
    log_1_5 = "The number of observations with missing specimen type: "
    log_1_6 = "The number of observations with missing culture result: "
    log_1_7 = "Example format of specimen date: "
    log_1_8 = "Format of microbiology data file: "
    log_1_9 = "Missing data− Hospital admission data set (file name: hospital_admission_data)"
    log_1_10 = "The number of observations with missing admission date: "
    log_1_11 = "The number of observations with missing discharge date: "
    log_1_12 = "The number of observations with missing outcome: "
    log_1_13 = "Example format of admission date: "
    log_1_14 = "Example format of discharge date: "
    log_1_15 = "Missing data− Merged data set (merged by the AMASS application)"
    log_1_16 = "The number of observations with missing specimen date: "
    log_1_17 = "The number of observations with missing admission date: "
    log_1_18 = "The number of observations with missing discharge date: "
    log_1_19 = "The number of observations with missing age: "
    log_1_20 = "The number of observations with missing gender: "
    log_1_21 = "The number of observations with missing infection origin data: "
    log_1 = ["<b>" + log_1_1 + "</b>", 
            add_blankline + log_1_2,
            add_blankline + "<b>" + log_1_3 + "</b>", 
            log_1_4 + assign_na_toinfo(retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_missing_specimen_date","Parameters")),
            log_1_5 + assign_na_toinfo(retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_missing_specimen_type","Parameters")),
            log_1_6 + assign_na_toinfo(retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_missing_culture_result","Parameters")),
            log_1_7 + assign_na_toinfo(retrieve_results(over_raw,"microbiology_data","Type_of_data_file","format_of_specimen_date","Parameters")),
            log_1_8 + str(file_format),
            add_blankline + "<b>" + log_1_9 + "</b>", 
            log_1_10 + assign_na_toinfo(retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_missing_admission_date","Parameters")),
            log_1_11 + assign_na_toinfo(retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_missing_discharge_type","Parameters")),
            log_1_12 + assign_na_toinfo(retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_missing_outcome_result","Parameters")),
            log_1_13 + assign_na_toinfo(retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","format_of_admission_date","Parameters")),
            log_1_14 + assign_na_toinfo(retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","format_of_discharge_date","Parameters")),
            add_blankline + "<b>" + log_1_15 + "</b>", 
            log_1_16 + assign_na_toinfo(retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_missing_specimen_date","Parameters")),
            log_1_17 + assign_na_toinfo(retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_missing_admission_date","Parameters")),
            log_1_18 + assign_na_toinfo(retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_missing_discharge_type","Parameters")),
            log_1_19 + assign_na_toinfo(retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_missing_age","Parameters")),
            log_1_20 + assign_na_toinfo(retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_missing_gender","Parameters")),
            log_1_21 + assign_na_toinfo(retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_missing_infection_origin_data","Parameters"))]
    report_title(c,'Data verification log file',1.07*inch, 10.5*inch,'#3e4444', font_size=16)
    report_context(c,log_1, 1.07*inch, 3.2*inch, 460, 500, font_size=11)
    c.showPage()

def ts(df, df_col, marked_idx, title_name, title_sub):
    if len(marked_idx) == 1: # no.row < 30 rows
        log_1 = "<b>" + title_name + ": " + title_sub + "</b>"
        df_1 = df.loc[0:]
        df_1 = df_1.values.tolist()
        df_1 = df_col + df_1
        report_context(c,[log_1], 1.0*inch, 10.0*inch, 460, 80, font_size=11)
        table = df_1
        table_draw = report_table_appendix(table)
        table_draw.wrapOn(c, 500, 300)
        h = (30-len(table))*(0.25) ####Work!!!!!!!!!!!!!!!!!!!!!!!!
        table_draw.drawOn(c, 1.07*inch, (h+2.0)*inch)
        c.showPage()
    else:                        # no.row >= 30 rows
        for i in range(len(marked_idx)):
            if i == 0:
                log_1 = "<b>" + title_name + ": " + title_sub + "</b>"
                df_1 = df.loc[marked_idx[i]:marked_idx[i+1]]
            elif i+1 == len(marked_idx):
                log_1 = "<b>" + title_name + " (continue): " + title_sub + "</b>"
                df_1 = df.loc[marked_idx[i]:]
            else:
                log_1 = "<b>" + title_name + " (continue): " + title_sub + "</b>"
                df_1 = df.loc[marked_idx[i]+1:marked_idx[i+1]]
            df_1 = df_1.values.tolist()
            df_1 = df_col + df_1
            report_context(c,[log_1], 1.0*inch, 10.0*inch, 460, 80, font_size=11)
            table = df_1
            table_draw = report_table_appendix(table)
            table_draw.wrapOn(c, 500, 300)
            h = (30-len(table))*(0.25) ####Work!!!!!!!!!!!!!!!!!!!!!!!!
            table_draw.drawOn(c, 1.07*inch, (h+2.0)*inch)
            c.showPage()

##Generating logfile_amass.pdf and Exporting logfile.xlsx
try:
    c = canvas.Canvas(path + "Data_verification_logfile_report.pdf")
    if checkpoint(over_i):
        try:
            cover(over_raw)
            ts0(over_raw)
        except Exception as e:
            logger.exception(e)
            pass
    if checkpoint(dict_i):
        if checkpoint(org_i):
            try:
                ts(org_A, [list(org_A.columns)], marked_org_A, "Table S1A", "List of data values of the variable recorded for \"organism\" in your microbiology data file, which are mainly used for the main report")
                org_A.columns = [w.replace("\n"," ") for w in org_A.columns.tolist()]
                org_A.to_excel(path + "ResultData/logfile_TS1A_main_organisms.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
            try:
                ts(org_B, [list(org_B.columns)], marked_org_B, "Table S1B", "List of data values of the variable recorded for \"organism\" in your microbiology data file, which are mainly used for the annex")
                org_B.columns = [w.replace("\n"," ") for w in org_B.columns.tolist()]
                org_B.to_excel(path + "ResultData/logfile_TS1B_optional_organisms.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
        if checkpoint(spc_i):
            try:
                ts(spc_merge, [list(spc_merge.columns)], marked_spc, "Table S2", "List of data values of the variable recorded for \"specimen_type\" in your microbiology data file")
                spc_merge.columns = [w.replace("\n"," ") for w in spc_merge.columns.tolist()]
                spc_merge.to_excel(path + "ResultData/logfile_TS2_specimens.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
        if checkpoint(ast_i):
            try:
                ts(ast_merge, [list(ast_merge.columns)], marked_ast, "Table S3", "List of variables recorded for \"antibiotics\" in your microbiology data file")
                ast_merge.columns = [w.replace("\n"," ") for w in ast_merge.columns.tolist()]
                ast_merge.to_excel(path + "ResultData/logfile_TS3_antibiotics.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
    if checkpoint(dict_hosp_i):
        if checkpoint(gen_i):
            try:
                ts(gen_merge, [list(gen_merge.columns)], marked_gen, "Table S4", "List of data values of variable recorded for \"gender\" in your hospital admission data file")
                gen_merge.columns = [w.replace("\n"," ") for w in gen_merge.columns.tolist()]
                gen_merge.to_excel(path + "ResultData/logfile_TS4_gender.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
        if checkpoint(age_i):
            try:
                ts(age_merge, [list(age_merge.columns)], marked_age, "Table S5", "List of data values of variable recorded for \"age\" in your hospital admission data file")
                age_merge.columns = [w.replace("\n"," ") for w in age_merge.columns.tolist()]
                age_merge.to_excel(path + "ResultData/logfile_TS5_age.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
        if checkpoint(dis_i):
            try:
                ts(dis_merge, [list(dis_merge.columns)], marked_dis, "Table S6", "List of data values of variable recorded for \"discharge status\" in your hospital admission data file")
                dis_merge.columns = [w.replace("\n"," ") for w in dis_merge.columns.tolist()]
                dis_merge.to_excel(path + "ResultData/logfile_TS6_discharge_status.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
    if checkpoint(var_mi_i):
        try:
            ts(var_mi_raw, var_mi_col, marked_var_mi, "Table S7", "List of variable names in your microbiology_data file")
        except Exception as e:
            logger.exception(e)
            pass
    if checkpoint(var_ho_i):
        try:
            ts(var_ho_raw, var_ho_col, marked_var_ho, "Table S8", "List of variable names in your hospital_admission_data file")
        except Exception as e:
            logger.exception(e)
            pass
    c.save()
except Exception as e:
    logger.exception(e)
    pass