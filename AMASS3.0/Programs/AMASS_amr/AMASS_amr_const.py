#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** CONST FILE and Configurations                                                                   ***#
#***-------------------------------------------------------------------------------------------------***#
# @author: PRAPASS WANNAPINIJ
# Created on: 09 MAR 2023 
import pandas as pd #for creating and manipulating dataframe

CONST_DICTCOL_AMASS = "amassvar"
CONST_DICTCOL_DATAVAL = "dataval"
CONST_DICVAL_EMPTY = ""
CONST_ORIGIN_DATE = pd.to_datetime("1899-12-30")
CONST_CDATEFORMAT = ["%d/%h/%y", "%d/%h/%Y","%d/%m/%y", "%d/%m/%Y","%d/%b/%y", "%d/%b/%Y","%d-%h-%y", "%d-%h-%Y","%d-%m-%y",  "%d-%m-%Y","%d-%b-%y",  "%d-%b-%Y","%d%h%y", "%d%h%Y","%d%m%y", "%d%m%Y","%d%b%y", "%d%b%Y","%Y-%m-%d"]
CONST_ORG_NOTINTEREST_ORGCAT = 0
CONST_ORG_NOGROWTH = "organism_no_growth"
CONST_ORG_NOGROWTH_ORGCAT = 9
CONST_ORG_NOGROWTH_ORGCAT_ANNEX_A = 99
CONST_MERGE_DUPCOLSUFFIX_HOSP = "_hosp"
CONST_YEARAGECAT_CUT_UNKNOWN_LABEL = "Unknown"
CONST_YEARAGECAT_CUT_VALUE = [0, 1, 5, 15, 25, 35, 45, 55, 65, 81, 200]
CONST_YEARAGECAT_CUT_LABEL = ["less than 1 year","1 to 4 years","5 to 14 years","15 to 24 years","25 to 34 years","35 to 44 years","45 to 54 years","55 to 64 years","65 to 80 years","over 80 years"]
CONST_DICT_COHO_CO = "community_origin"
CONST_DICT_COHO_HO = "hospital_origin"
CONST_DICT_COHO_AVAIL = "infection_origin_available"
CONST_EXPORT_COHO_CO_DATAVAL = "Community"
CONST_EXPORT_COHO_HO_DATAVAL = "Hospital"
CONST_EXPORT_COHO_MORTALITY_CO_DATAVAL = "Community-origin"
CONST_EXPORT_COHO_MORTALITY_HO_DATAVAL = "Hospital-origin"
CONST_EXPORT_MORTALITY_ATB_R_SUFFIX = "-NS"
CONST_EXPORT_MORTALITY_ATB_I_SUFFIX = "-S"
CONST_EXPORT_MORTALITY_COHO_SUFFIX = "-origin"
CONST_DIED_VALUE = "died"
CONST_ALIVE_VALUE = "alive"
CONST_PERPOP = 100000
CONST_PATH_RESULT = "./ResultData/"
CONST_PATH_VAR = "./Variables/"
CONST_PATH_ROOT = "./"
CONST_ANNEXA_NON_ORG = "non-organism-annex-A"

#Const variable name in data dictionary
CONST_VARNAME_HOSPITALNUMBER = "hospital_number"
CONST_VARNAME_ADMISSIONDATE = "date_of_admission"
CONST_VARNAME_DISCHARGEDATE = "date_of_discharge"
CONST_VARNAME_DISCHARGESTATUS = "discharge_status"
CONST_VARNAME_GENDER = "gender"
CONST_VARNAME_BIRTHDAY = "birthday"
CONST_VARNAME_AGEY = "age_year"
CONST_VARNAME_AGEGROUP = "age_group"

CONST_VARNAME_SPECDATERAW = "specimen_collection_date"
CONST_VARNAME_COHO = "infection_origin"
CONST_VARNAME_ORG = "organism"
CONST_VARNAME_SPECTYPE = "specimen_type"


# COnst for new column hospital data
CONST_NEWVARNAME_HN_HOSP ="hn" + CONST_MERGE_DUPCOLSUFFIX_HOSP
CONST_NEWVARNAME_DAYTOADMDATE = "admdate2"
CONST_NEWVARNAME_CLEANADMDATE = "DateAdm"
CONST_NEWVARNAME_ADMYEAR = "YearAdm"
CONST_NEWVARNAME_ADMMONTHNAME = "MonthAdm"
CONST_NEWVARNAME_DAYTODISDATE = "disdate2"
CONST_NEWVARNAME_CLEANDISDATE = "DateDis"
CONST_NEWVARNAME_DISYEAR = "YearDis"
CONST_NEWVARNAME_DISMONTHNAME = "MonthDis"
CONST_NEWVARNAME_DAYTOBIRTHDATE = "bdate2"
CONST_NEWVARNAME_CLEANBIRTHDATE = "DateBirth"
CONST_NEWVARNAME_DISOUTCOME_HOSP = "disoutcome2_hosp"
CONST_NEWVARNAME_DISOUTCOME = "disoutcome2"
CONST_NEWVARNAME_PATIENTDAY= "patientdays"
CONST_NEWVARNAME_PATIENTDAY_HO= "patientdays_his"
# Const for new column micro
CONST_NEWVARNAME_BLOOD ="blood"
CONST_NEWVARNAME_ORG3 ="organism3"
CONST_NEWVARNAME_HN ="hn"
CONST_NEWVARNAME_DAYTOSPECDATE = "spcdate2"
CONST_NEWVARNAME_CLEANSPECDATE = "DateSpc"
CONST_NEWVARNAME_SPECYEAR = "YearSpc"
CONST_NEWVARNAME_SPECMONTHNAME = "MonthSpc"
CONST_NEWVARNAME_ORGCAT = "organismCat"
CONST_NEWVARNAME_PREFIX_AST = "AST"
CONST_NEWVARNAME_PREFIX_RIS = "RIS"
CONST_NEWVARNAME_AST3GC = CONST_NEWVARNAME_PREFIX_AST + "3gc"
CONST_NEWVARNAME_ASTCBPN = CONST_NEWVARNAME_PREFIX_AST + "Carbapenem"
CONST_NEWVARNAME_ASTFRQ = CONST_NEWVARNAME_PREFIX_AST + "Fluoroquin"
CONST_NEWVARNAME_ASTTETRA = CONST_NEWVARNAME_PREFIX_AST + "Tetra"
CONST_NEWVARNAME_ASTAMINOGLY = CONST_NEWVARNAME_PREFIX_AST + "aminogly"
CONST_NEWVARNAME_ASTMRSA = CONST_NEWVARNAME_PREFIX_AST + "mrsa"
CONST_NEWVARNAME_ASTPEN = CONST_NEWVARNAME_PREFIX_AST + "pengroup"
CONST_NEWVARNAME_AST3GCCBPN = CONST_NEWVARNAME_PREFIX_AST + "3gcsCarbs"
CONST_NEWVARNAME_AMR = "AMR"
CONST_NEWVARNAME_AMRCAT = "AMRcat"
CONST_NEWVARNAME_MICROREC_ID ="amass_micro_rec_id"
CONST_NEWVARNAME_MATCHLEV ="hospmicro_match_level"
# Const for new column in merge hosp/micro data
CONST_NEWVARNAME_GENDERCAT ="gender_cat"
CONST_NEWVARNAME_AGEYEAR ="YearAge"
CONST_NEWVARNAME_AGECAT ="YearAge_cat"
CONST_NEWVARNAME_DAYADMTOSPC = "losSpc"
CONST_NEWVARNAME_COHO_FROMHOS = "InfOri_hosp"
CONST_NEWVARNAME_COHO_FROMCAL = "InfOri_cal"
CONST_NEWVARNAME_COHO_FINAL = "InfOri"
# COnst for new column in ANNEX A
CONST_NEWVARNAME_SPECTYPE_ANNEXA = "spectype_A"
CONST_NEWVARNAME_SPECTYPENAME_ANNEXA = "spectypename_A"
CONST_NEWVARNAME_ORG3_ANNEXA ="organismA"
CONST_NEWVARNAME_ORGCAT_ANNEXA = "organismCat_A"
CONST_NEWVARNAME_ORGNAME_ANNEXA = "organismname_A"
def dict_ris(df_dict) :
    dict_ris_temp = {}
    temp_df = df_dict[df_dict[CONST_DICTCOL_AMASS]=="resistant"]
    for index, row in temp_df .iterrows():
        dict_ris_temp.update({row[CONST_DICTCOL_DATAVAL]: "R"})
    temp_df = df_dict[df_dict[CONST_DICTCOL_AMASS]=="intermediate"]
    for index, row in temp_df .iterrows():
        dict_ris_temp.update({row[CONST_DICTCOL_DATAVAL]: "I"})
    temp_df = df_dict[df_dict[CONST_DICTCOL_AMASS]=="susceptible"]
    for index, row in temp_df .iterrows():
        dict_ris_temp.update({row[CONST_DICTCOL_DATAVAL]: "S"})
    return dict_ris_temp
dict_ast = {
    "R":"1",
    "I":"1",
    "S":"0"
}
def dict_orgcatwithatb(bisabom):
    return {
    "organism_staphylococcus_aureus":[1,1,"Staphylococcus aureus",
                                      ["Methicillin","Vancomycin","Clindamycin"],
                                      [CONST_NEWVARNAME_ASTMRSA,"RISVancomycin","RISClindamycin"]],
    "organism_enterococcus_spp":[2,1,"Enterococcus spp.",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin"]],
    "organism_streptococcus_pneumoniae":[7,1,"Streptococcus pneumoniae",
                                         ["Penicillin G","Oxacillin","Co-trimoxazole","3GC","Ceftriaxone","Cefotaxime","Erythromycin","Clindamycin","Levofloxacin"],
                                         ["RISPenicillin_G","RISOxacillin","RISSulfamethoxazole_and_trimethoprim",CONST_NEWVARNAME_AST3GC,"RISCeftriaxone","RISCefotaxime","RISErythromycin","RISClindamycin","RISLevofloxacin"]],
    "organism_salmonella_spp":[8,1,"Salmonella spp.",
                                      ["FLUOROQUINOLONES","Ciprofloxacin","Levofloxacin","3GC","Ceftriaxone","Cefotaxime","Ceftazidime","CARBAPENEMS","Imipenem","Meropenem","Ertapenem","Doripenem"],
                                      [CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCeftriaxone","RISCefotaxime","RISCeftazidime",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem"]],
    "organism_escherichia_coli":[3,1,"Escherichia coli",
                                      ["Gentamicin","Amikacin","Co-trimoxazole","Ampicillin","FLUOROQUINOLONES","Ciprofloxacin","Levofloxacin","3GC","Cefpodoxime","Ceftriaxone","Cefotaxime","Ceftazidime",
                                       "Cefepime","CARBAPENEMS","Imipenem","Meropenem","Ertapenem","Doripenem","Colistin"],
                                      ["RISGentamicin","RISAmikacin","RISSulfamethoxazole_and_trimethoprim","RISAmpicillin",CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCefpodoxime","RISCeftriaxone","RISCefotaxime","RISCeftazidime",
                                       "RISCefepime",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem","RISColistin"]],
    "organism_klebsiella_pneumoniae":[4,1,"Klebsiella pneumoniae",
                                      ["Gentamicin","Amikacin","Co-trimoxazole","FLUOROQUINOLONES","Ciprofloxacin","Levofloxacin","3GC","Cefpodoxime","Ceftriaxone","Cefotaxime","Ceftazidime",
                                       "Cefepime","CARBAPENEMS","Imipenem","Meropenem","Ertapenem","Doripenem","Colistin"],
                                      ["RISGentamicin","RISAmikacin","RISSulfamethoxazole_and_trimethoprim",CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCefpodoxime","RISCeftriaxone","RISCefotaxime","RISCeftazidime",
                                       "RISCefepime",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem","RISColistin"]],
    "organism_pseudomonas_aeruginosa":[5,1,"Pseudomonas aeruginosa",
                                      ["Ceftazidime","Ciprofloxacin","Piperacillin/tazobactam","AMINOGLYCOSIDES",
                                       "Gentamicin","Amikacin","CARBAPENEMS","Imipenem","Meropenem","Doripenem","Colistin"],
                                      ["RISCeftazidime","RISCiprofloxacin","RISPiperacillin_and_tazobactam",CONST_NEWVARNAME_ASTAMINOGLY,
                                       "RISGentamicin","RISAmikacin",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISDoripenem","RISColistin"]],
    "organism_acinetobacter_baumannii" if bisabom==True else "organism_acinetobacter_spp":[6,1,"Acinetobacter baumannii" if bisabom==True else "Acinetobacter spp.",
                                      ["Tigecycline","Minocycline","AMINOGLYCOSIDES","Gentamicin","Amikacin",
                                       "CARBAPENEMS","Imipenem","Meropenem","Doripenem","Colistin"],
                                      ["RISTigecycline","RISMinocycline",CONST_NEWVARNAME_ASTAMINOGLY,"RISGentamicin","RISAmikacin",
                                       CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISDoripenem","RISColistin"]],
    CONST_ORG_NOGROWTH:[CONST_ORG_NOGROWTH_ORGCAT,0,"No growth",[],[]]
    }
def dict_orgwithatb_mortality(bisabom):
    return {"organism_staphylococcus_aureus":["Staphylococcus aureus",["MRSA","MSSA"],
                                                                [CONST_NEWVARNAME_ASTMRSA,CONST_NEWVARNAME_ASTMRSA],["1","0"]],
                             "organism_enterococcus_spp":["Enterococcus spp.",["Vancomycin-NS","Vancomycin-S"],
                                                                ["ASTVancomycin","ASTVancomycin"],["1","0"]],
                             "organism_streptococcus_pneumoniae":["Streptococcus pneumoniae",["Penicillin-NS","Penicillin-S"],
                                                                ["ASTPenicillin_G","ASTPenicillin_G"],["1","0"]],
                             "organism_salmonella_spp":["Salmonella spp.",["Fluoroquinolone-NS","Fluoroquinolone-S"],
                                                                [CONST_NEWVARNAME_ASTFRQ,CONST_NEWVARNAME_ASTFRQ],["1","0"]],
                             "organism_escherichia_coli":["Escherichia coli",["Carbapenem-NS","3GC-NS","3GC-S"],
                                                                [CONST_NEWVARNAME_ASTCBPN,CONST_NEWVARNAME_AST3GCCBPN,CONST_NEWVARNAME_AST3GCCBPN],["1","2","1"]],
                             "organism_klebsiella_pneumoniae":["Klebsiella pneumoniae",["Carbapenem-NS","3GC-NS","3GC-S"],
                                                                [CONST_NEWVARNAME_ASTCBPN,CONST_NEWVARNAME_AST3GCCBPN,CONST_NEWVARNAME_AST3GCCBPN],["1","2","1"]],
                             "organism_pseudomonas_aeruginosa":["Pseudomonas aeruginosa",["Carbapenem-NS","Carbapenem-S"],
                                                                [CONST_NEWVARNAME_ASTCBPN,CONST_NEWVARNAME_ASTCBPN],["1","0"]],
                             "organism_acinetobacter_baumannii" if bisabom==True else "organism_acinetobacter_spp":["Acinetobacter baumannii" if bisabom==True else "Acinetobacter spp.",["Carbapenem-NS","Carbapenem-S"],
                                                                [CONST_NEWVARNAME_ASTCBPN,CONST_NEWVARNAME_ASTCBPN],["1","0"]]
                          } 
def dict_orgwithatb_incidence(bisabom):
    return {"organism_staphylococcus_aureus":["S. aureus",["MRSA"],[CONST_NEWVARNAME_ASTMRSA],["1"]],
                             "organism_enterococcus_spp":["Enterococcus spp.",["Vancomycin-NSEnterococcus spp."],["ASTVancomycin"],["1"]],
                             "organism_streptococcus_pneumoniae":["S. pneumoniae",["Penicillin-NSS. pneumoniae"],["ASTPenicillin_G"],["1"]],
                             "organism_salmonella_spp":["Salmonella spp.",["Fluoroquinolone-NSSalmonella spp."],[CONST_NEWVARNAME_ASTFRQ],["1"]],
                             "organism_escherichia_coli":["E. coli",["3GC-NS\nE. coli","Carbapenem-NSE. coli"],[CONST_NEWVARNAME_AST3GC,CONST_NEWVARNAME_ASTCBPN],["1","1"]],
                             "organism_klebsiella_pneumoniae":["K. pneumoniae",["3GC-NSK. pneumoniae","Carbapenem-NS\nK. pneumoniae"],[CONST_NEWVARNAME_AST3GC,CONST_NEWVARNAME_ASTCBPN],["1","1"]],
                             "organism_pseudomonas_aeruginosa":["P. aeruginosa",["Carbapenem-NSP. aeruginosa"],[CONST_NEWVARNAME_ASTCBPN],["1"]],
                             "organism_acinetobacter_baumannii" if bisabom==True else "organism_acinetobacter_spp":["A. baumannii" if bisabom==True else "Acinetobacter spp.",["Carbapenem-NSA. baumannii" if bisabom==True else "Carbapenem-NSAcinetobacter spp."],[CONST_NEWVARNAME_ASTCBPN],["1"]]
                          } 
# We can convert antibiotic list to read from configuration file that we can configure atb vs organism to set it RIS level that will start consider as AST=1 (Currently I and S -> AST =1)
list_antibiotic = ["Amikacin","Amoxicillin","Amoxicillin_and_clavulanic_acid","Ampicillin","Ampicillin_and_sulbactam","Aztreonam","Cefazolin",
                 "Cefepime","Cefotaxime","Cefotetan","Cefoxitin","Cefpodoxime","Ceftaroline",
                 "Ceftazidime","Ceftriaxone","Cefuroxime","Chloramphenicol","Ciprofloxacin",
                 "Clarithromycin","Clindamycin","Colistin","Dal fopristin_and_quinupristin",
                 "Daptomycin","Doripenem","Doxycycline","Ertapenem","Erythromycin","Fosfomycin",
                 "Fusidic_acid","Gentamicin","Imipenem","Levofloxacin","Linezolid","Meropenem","Methicillin","Minocycline",
                 "Moxifloxacin","Nalidixic_acid","Netilmicin","Nitrofurantoin","Oxacillin","Penicillin_G","Piperacillin_and_tazobactam",
                 "Polymyxin_B","Rifampin","Streptomycin","Teicoplanin","Telavancin","Tetracycline","Ticarcillin_and_clavulanic_acid","Tigecycline","Tobramycin","Trimethoprim","Sulfamethoxazole_and_trimethoprim","Vancomycin"]
