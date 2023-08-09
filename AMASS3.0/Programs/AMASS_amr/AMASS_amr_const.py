#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** CONST FILE and Configurations                                                                   ***#
#***-------------------------------------------------------------------------------------------------***#
# @author: PRAPASS WANNAPINIJ
# Created on: 09 MAR 2023 
import pandas as pd #for creating and manipulating dataframe
CONST_SOFTWARE_VERSION ="3.0 Build 3004 on 8 August 2023"

CONST_DIR_INPUT = "./"
CONST_DIR_RESULTDATA = "./ResultData/"
CONST_SPECIFY_CODE_AS_NOTUSEDORG = 99
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
CONST_PATH_REPORTWITH_PID = "./Report_with_patient_identifiers/"
CONST_PATH_ROOT = "./"
CONST_ANNEXA_NON_ORG = "non-organism-annex-A"
# COnst for ANNEX B using data after map with dict or original microbiology data file.
CONST_ANNEXB_USING_MAPPEDDATA = False
#ATB name
CONST_ATBNAME_AST3GC = "3GC"
CONST_ATBNAME_ASTCBPN = "CARBAPENEMS"
CONST_ATBNAME_ASTFRQ = "FLUOROQUINOLONES"
CONST_ATBNAME_ASTTETRA = "Tetra"
CONST_ATBNAME_ASTAMINOGLY = "AMINOGLYCOSIDES"
CONST_ATBNAME_ASTMRSA = "Methicillin"
CONST_ATBNAME_ASTPEN = "pengroup"
CONST_ATBNAME_AST3GCCBPN = "3GC CARBAPENEMS"
#For generate graph in report
CONST_MAXNUMCHAR_ATBNAME = 18
CONST_BARGRAPH_ITEM_HEIGHT_INCH = 0.5
CONST_BARGRAPH_OTHERPART_HEIGHT_INCH = 1.5


#Const variable name in data dictionary
CONST_VARNAME_HOSPITALNUMBER = "hospital_number"
CONST_VARNAME_ADMISSIONDATE = "date_of_admission"
CONST_VARNAME_DISCHARGEDATE = "date_of_discharge"
CONST_VARNAME_DISCHARGESTATUS = "discharge_status"
CONST_VARNAME_GENDER = "gender"
CONST_VARNAME_BIRTHDAY = "birthday"
CONST_VARNAME_AGEY = "age_year"
CONST_VARNAME_AGEGROUP = "age_group"
CONST_VARNAME_SPECNUM = "specimen_number"
CONST_VARNAME_WARD = "ward"

CONST_VARNAME_SPECDATERAW = "specimen_collection_date"
CONST_VARNAME_SPECRPTDATERAW = "specimen_result_report_date"
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
CONST_NEWVARNAME_DAYTOSTARTDATE = "startdate2"
CONST_NEWVARNAME_DAYTOENDDATE = "enddate2"
CONST_NEWVARNAME_PATIENTDAY= "patientdays"
CONST_NEWVARNAME_PATIENTDAY_HO= "patientdays_his"
CONST_NEWVARNAME_WARDCODE= "ward_code"
CONST_NEWVARNAME_WARDTYPE= "ward_type"
# Const for new column micro
CONST_NEWVARNAME_AMASSSPECTYPE ="amass_spectype"
CONST_NEWVARNAME_BLOOD ="blood"
CONST_NEWVARNAME_ORG3 ="organism3"
CONST_NEWVARNAME_HN ="hn"
CONST_NEWVARNAME_DAYTOSPECDATE = "spcdate2"
CONST_NEWVARNAME_CLEANSPECDATE = "DateSpc"
CONST_NEWVARNAME_SPECYEAR = "YearSpc"
CONST_NEWVARNAME_SPECMONTHNAME = "MonthSpc"
CONST_NEWVARNAME_ORGCAT = "organismCat"
CONST_NEWVARNAME_PREFIX_AST = "NS_"
CONST_NEWVARNAME_PREFIX_RIS = "RIS"
CONST_NEWVARNAME_DAYTOSPECRPTDATE = "spcrptdate2"
CONST_NEWVARNAME_CLEANSPECRPTDATE = "DateSpcRpt"
CONST_NEWVARNAME_SPECRPTYEAR = "YearSpcRpt"
CONST_NEWVARNAME_SPECRPTMONTHNAME = "MonthSpcRpt"
#Antibiotic group RIS value
CONST_NEWVARNAME_AST3GC_RIS = CONST_NEWVARNAME_PREFIX_RIS + "3gc"
CONST_NEWVARNAME_ASTCBPN_RIS = CONST_NEWVARNAME_PREFIX_RIS + "Carbapenem"
CONST_NEWVARNAME_ASTFRQ_RIS = CONST_NEWVARNAME_PREFIX_RIS + "Fluoroquin"
CONST_NEWVARNAME_ASTTETRA_RIS = CONST_NEWVARNAME_PREFIX_RIS + "Tetra"
CONST_NEWVARNAME_ASTAMINOGLY_RIS = CONST_NEWVARNAME_PREFIX_RIS + "aminogly"
CONST_NEWVARNAME_ASTMRSA_RIS = CONST_NEWVARNAME_PREFIX_RIS + "mrsa"
CONST_NEWVARNAME_ASTPEN_RIS = CONST_NEWVARNAME_PREFIX_RIS + "pengroup"
CONST_NEWVARNAME_AST3GCCBPN_RIS = CONST_NEWVARNAME_PREFIX_RIS + "3gcsCarbs"
#Antibiotic group AST value
CONST_NEWVARNAME_AST3GC = CONST_NEWVARNAME_PREFIX_AST + "3gc"
CONST_NEWVARNAME_ASTCBPN = CONST_NEWVARNAME_PREFIX_AST + "Carbapenem"
CONST_NEWVARNAME_ASTFRQ = CONST_NEWVARNAME_PREFIX_AST + "Fluoroquin"
CONST_NEWVARNAME_ASTTETRA = CONST_NEWVARNAME_PREFIX_AST + "Tetra"
CONST_NEWVARNAME_ASTAMINOGLY = CONST_NEWVARNAME_PREFIX_AST + "aminogly"
CONST_NEWVARNAME_ASTMRSA = CONST_NEWVARNAME_PREFIX_AST + "mrsa"
CONST_NEWVARNAME_ASTPEN = CONST_NEWVARNAME_PREFIX_AST + "pengroup"
CONST_NEWVARNAME_AST3GCCBPN = CONST_NEWVARNAME_PREFIX_AST + "3gcsCarbs"
CONST_NEWVARNAME_AMR = "AMR"
CONST_NEWVARNAME_AMR_TESTED = "AMR_TESTED"
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

#Mode for display only R or AST value
CONST_VALUE_MODE_ONLYR = 1
CONST_VALUE_MODE_AST = 2
CONST_MODE_R_OR_AST = CONST_VALUE_MODE_AST

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
def dict_ast() :
    return {
        "R":"1",
        "I":"1" if CONST_MODE_R_OR_AST == CONST_VALUE_MODE_AST else "0",
        "S":"0"
    }
"""
dict_ast = {
    "R":"1",
    "I":"1",
    "S":"0"
}
"""
def getlist_amr_atb(dict_orgatb):
    try:
        list_atb = []
        for sorgkey in dict_orgatb:
            ocurorg = dict_orgatb[sorgkey]
            if ocurorg[1] == 1 :
                list_Curatb_source = ocurorg[4]
                list_curatb = []
                for atb in list_Curatb_source:
                    if atb[0:len(CONST_NEWVARNAME_PREFIX_RIS)] == CONST_NEWVARNAME_PREFIX_RIS:
                        list_curatb = list_curatb + [CONST_NEWVARNAME_PREFIX_AST +  atb[len(CONST_NEWVARNAME_PREFIX_RIS):]]
                    else:
                        list_curatb = list_curatb + [atb]
                list_atb = list_atb + list_curatb
        #Dedup
        list_atb = list(dict.fromkeys(list_atb))
        return list_atb
    except Exception:
        return []
def dict_orgcatwithatb(bisabom,bisentspp):
    dict_org = get_dict_orgcatwithatb(bisabom,bisentspp)
    lst_delkey = []
    for skey in dict_org:
        lval = dict_org[skey]
        if lval[1] > 1:
            lst_delkey.append(skey)
    for skey in lst_delkey:
        dict_org.pop(skey)
    return dict_org
def get_dict_orgcatwithatb(bisabom,bisentspp): #Last line of antibiotic list is antibiotic added in version 3.0
    return {
    "organism_staphylococcus_aureus":[10,1,"Staphylococcus aureus",
                                      [CONST_ATBNAME_ASTMRSA,"Vancomycin","Clindamycin","Cefoxitin",
                                       "Oxacillin"],
                                      [CONST_NEWVARNAME_ASTMRSA_RIS,"RISVancomycin","RISClindamycin","RISCefoxitin"
                                       ,"RISOxacillin"],
                                      "<i>Staphylococcus aureus</i>"],
    "organism_enterococcus_spp":[20,1 if bisentspp==True else CONST_SPECIFY_CODE_AS_NOTUSEDORG,"Enterococcus spp.",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin"],
                                 "<i>Enterococcus</i> spp."],
    "organism_enterococcus_faecalis":[21,CONST_SPECIFY_CODE_AS_NOTUSEDORG if bisentspp==True else 1,"Enterococcus Faecalis",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin",
                                  "Gentamicin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin",
                                  "RISGentamicin"],
                                 "<i>Enterococcus Faecalis</i>"],
    "organism_enterococcus_faecium":[22,CONST_SPECIFY_CODE_AS_NOTUSEDORG if bisentspp==True else 1,"Enterococcus Faecium",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin",
                                  "Gentamicin","Amoxicillin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin",
                                  "RISGentamicin","RISAmoxicillin"],
                                 "<i>Enterococcus Faecium</i>"],
    "organism_streptococcus_pneumoniae":[70,1,"Streptococcus pneumoniae",
                                         ["Penicillin G","Oxacillin","Co-trimoxazole",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Erythromycin","Clindamycin","Levofloxacin"],
                                         ["RISPenicillin_G","RISOxacillin","RISSulfamethoxazole_and_trimethoprim",CONST_NEWVARNAME_AST3GC_RIS,"RISCeftriaxone","RISCefotaxime","RISErythromycin","RISClindamycin","RISLevofloxacin"],
                                         "<i>Streptococcus pneumoniae</i>"],
    "organism_salmonella_spp":[80,1,"Salmonella spp.",
                                      [CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Ceftazidime","Cefixime",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Ertapenem","Doripenem"],
                                      [CONST_NEWVARNAME_ASTFRQ_RIS,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC_RIS,"RISCeftriaxone","RISCefotaxime","RISCeftazidime","RISCefixime",CONST_NEWVARNAME_ASTCBPN_RIS,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem"],
                                      "<i>Salmonella</i> spp."],
    "organism_escherichia_coli":[30,1,"Escherichia coli",
                                      ["Gentamicin","Amikacin","Co-trimoxazole","Ampicillin",CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Cefpodoxime","Ceftriaxone","Cefotaxime","Ceftazidime",
                                       "Cefepime",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Ertapenem","Doripenem","Colistin",
                                       "Piperacillin/tazobactam","cefoperazone/sulbactam"],
                                      ["RISGentamicin","RISAmikacin","RISSulfamethoxazole_and_trimethoprim","RISAmpicillin",CONST_NEWVARNAME_ASTFRQ_RIS,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC_RIS,"RISCefpodoxime","RISCeftriaxone","RISCefotaxime","RISCeftazidime",
                                       "RISCefepime",CONST_NEWVARNAME_ASTCBPN_RIS,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem","RISColistin",
                                       "RISPiperacillin_and_tazobactam","RISCefoperazone_and_sulbactam"],
                                      "<i>Escherichia coli</i>"],
    "organism_klebsiella_pneumoniae":[40,1,"Klebsiella pneumoniae",
                                      ["Gentamicin","Amikacin","Co-trimoxazole","Cefpodoxime",CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Ceftazidime","Cefixime",
                                       "Cefepime",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Ertapenem","Doripenem","Colistin",
                                       "Piperacillin/tazobactam","cefoperazone/sulbactam"],
                                      ["RISGentamicin","RISAmikacin","RISSulfamethoxazole_and_trimethoprim","RISCefpodoxime",CONST_NEWVARNAME_ASTFRQ_RIS,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC_RIS,"RISCeftriaxone","RISCefotaxime","RISCeftazidime","RISCefixime",
                                       "RISCefepime",CONST_NEWVARNAME_ASTCBPN_RIS,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem","RISColistin",
                                       "RISPiperacillin_and_tazobactam","RISCefoperazone_and_sulbactam"],
                                      "<i>Klebsiella pneumoniae</i>"],
    "organism_pseudomonas_aeruginosa":[50,1,"Pseudomonas aeruginosa",
                                      ["Ceftazidime","Ciprofloxacin","Piperacillin/tazobactam",CONST_ATBNAME_ASTAMINOGLY,
                                       "Gentamicin","Amikacin",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Doripenem","Colistin",
                                       "Levofloxacin","Tobramycin"],
                                      ["RISCeftazidime","RISCiprofloxacin","RISPiperacillin_and_tazobactam",CONST_NEWVARNAME_ASTAMINOGLY_RIS,
                                       "RISGentamicin","RISAmikacin",CONST_NEWVARNAME_ASTCBPN_RIS,"RISImipenem","RISMeropenem","RISDoripenem","RISColistin",
                                       "RISLevofloxacin","RISTobramycin"],
                                      "<i>Pseudomonas aeruginosa</i>"],
    "organism_acinetobacter_baumannii" if bisabom==True else "organism_acinetobacter_spp":[60,1,"Acinetobacter baumannii" if bisabom==True else "Acinetobacter spp.",
                                      ["Tigecycline","Minocycline",CONST_ATBNAME_ASTAMINOGLY,"Gentamicin","Amikacin",
                                       CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Doripenem","Colistin",
                                       CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Ceftazidime","Cefepime","Piperacillin/tazobactam","cefoperazone/sulbactam"],
                                      ["RISTigecycline","RISMinocycline",CONST_NEWVARNAME_ASTAMINOGLY_RIS,"RISGentamicin","RISAmikacin",
                                       CONST_NEWVARNAME_ASTCBPN_RIS,"RISImipenem","RISMeropenem","RISDoripenem","RISColistin",
                                       CONST_NEWVARNAME_ASTFRQ_RIS,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC_RIS,"RISCeftriaxone","RISCefotaxime","RISCeftazidime","RISCefepime","RISPiperacillin_and_tazobactam","RISCefoperazone_and_sulbactam"],
                                      "<i>Acinetobacter baumannii</i>" if bisabom==True else "<i>Acinetobacter</i> spp."],
    CONST_ORG_NOGROWTH:[CONST_ORG_NOGROWTH_ORGCAT,0,"No growth",[],[],"No growth"]
    }
"""
def get_dict_orgcatwithatb_V3old(bisabom,bisentspp): #Last line of antibiotic list is antibiotic added in version 3.0
    return {
    "organism_staphylococcus_aureus":[10,1,"Staphylococcus aureus",
                                      [CONST_ATBNAME_ASTMRSA,"Vancomycin","Clindamycin","Cefoxitin",
                                       "Oxacillin"],
                                      [CONST_NEWVARNAME_ASTMRSA,"RISVancomycin","RISClindamycin","RISCefoxitin"
                                       ,"RISOxacillin"],
                                      "<i>Staphylococcus aureus</i>"],
    "organism_enterococcus_spp":[20,1 if bisentspp==True else CONST_SPECIFY_CODE_AS_NOTUSEDORG,"Enterococcus spp.",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin"],
                                 "<i>Enterococcus</i> spp."],
    "organism_enterococcus_faecalis":[21,CONST_SPECIFY_CODE_AS_NOTUSEDORG if bisentspp==True else 1,"Enterococcus Faecalis",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin",
                                  "Gentamicin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin",
                                  "RISGentamicin"],
                                 "<i>Enterococcus Faecalis</i>"],
    "organism_enterococcus_faecium":[22,CONST_SPECIFY_CODE_AS_NOTUSEDORG if bisentspp==True else 1,"Enterococcus Faecium",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin",
                                  "Gentamicin","Amoxicillin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin",
                                  "RISGentamicin","RISAmoxicillin"],
                                 "<i>Enterococcus Faecium</i>"],
    "organism_streptococcus_pneumoniae":[70,1,"Streptococcus pneumoniae",
                                         ["Penicillin G","Oxacillin","Co-trimoxazole",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Erythromycin","Clindamycin","Levofloxacin"],
                                         ["RISPenicillin_G","RISOxacillin","RISSulfamethoxazole_and_trimethoprim",CONST_NEWVARNAME_AST3GC,"RISCeftriaxone","RISCefotaxime","RISErythromycin","RISClindamycin","RISLevofloxacin"],
                                         "<i>Streptococcus pneumoniae</i>"],
    "organism_salmonella_spp":[80,1,"Salmonella spp.",
                                      [CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Ceftazidime","Cefixime",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Ertapenem","Doripenem"],
                                      [CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCeftriaxone","RISCefotaxime","RISCeftazidime","RISCefixime",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem"],
                                      "<i>Salmonella</i> spp."],
    "organism_escherichia_coli":[30,1,"Escherichia coli",
                                      ["Gentamicin","Amikacin","Co-trimoxazole","Ampicillin",CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Cefpodoxime","Ceftriaxone","Cefotaxime","Ceftazidime",
                                       "Cefepime",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Ertapenem","Doripenem","Colistin",
                                       "Piperacillin/tazobactam","cefoperazone/sulbactam"],
                                      ["RISGentamicin","RISAmikacin","RISSulfamethoxazole_and_trimethoprim","RISAmpicillin",CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCefpodoxime","RISCeftriaxone","RISCefotaxime","RISCeftazidime",
                                       "RISCefepime",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem","RISColistin",
                                       "RISPiperacillin_and_tazobactam","RISCefoperazone_and_sulbactam"],
                                      "<i>Escherichia coli</i>"],
    "organism_klebsiella_pneumoniae":[40,1,"Klebsiella pneumoniae",
                                      ["Gentamicin","Amikacin","Co-trimoxazole","Cefpodoxime",CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Ceftazidime","Cefixime",
                                       "Cefepime",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Ertapenem","Doripenem","Colistin",
                                       "Piperacillin/tazobactam","cefoperazone/sulbactam"],
                                      ["RISGentamicin","RISAmikacin","RISSulfamethoxazole_and_trimethoprim","RISCefpodoxime",CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCeftriaxone","RISCefotaxime","RISCeftazidime","RISCefixime",
                                       "RISCefepime",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem","RISColistin",
                                       "RISPiperacillin_and_tazobactam","RISCefoperazone_and_sulbactam"],
                                      "<i>Klebsiella pneumoniae</i>"],
    "organism_pseudomonas_aeruginosa":[50,1,"Pseudomonas aeruginosa",
                                      ["Ceftazidime","Ciprofloxacin","Piperacillin/tazobactam",CONST_ATBNAME_ASTAMINOGLY,
                                       "Gentamicin","Amikacin",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Doripenem","Colistin",
                                       "Levofloxacin","Tobramycin"],
                                      ["RISCeftazidime","RISCiprofloxacin","RISPiperacillin_and_tazobactam",CONST_NEWVARNAME_ASTAMINOGLY,
                                       "RISGentamicin","RISAmikacin",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISDoripenem","RISColistin",
                                       "RISLevofloxacin","RISTobramycin"],
                                      "<i>Pseudomonas aeruginosa</i>"],
    "organism_acinetobacter_baumannii" if bisabom==True else "organism_acinetobacter_spp":[60,1,"Acinetobacter baumannii" if bisabom==True else "Acinetobacter spp.",
                                      ["Tigecycline","Minocycline",CONST_ATBNAME_ASTAMINOGLY,"Gentamicin","Amikacin",
                                       CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Doripenem","Colistin",
                                       CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Ceftazidime","Cefepime","Piperacillin/tazobactam","cefoperazone/sulbactam"],
                                      ["RISTigecycline","RISMinocycline",CONST_NEWVARNAME_ASTAMINOGLY,"RISGentamicin","RISAmikacin",
                                       CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISDoripenem","RISColistin",
                                       CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCeftriaxone","RISCefotaxime","RISCeftazidime","RISCefepime","RISPiperacillin_and_tazobactam","RISCefoperazone_and_sulbactam"],
                                      "<i>Acinetobacter baumannii</i>" if bisabom==True else "<i>Acinetobacter</i> spp."],
    CONST_ORG_NOGROWTH:[CONST_ORG_NOGROWTH_ORGCAT,0,"No growth",[],[],"No growth"]
    }
"""
"""
def get_dict_orgcatwithatb_V2(bisabom,bisentspp):
    return {
    "organism_staphylococcus_aureus":[10,1,"Staphylococcus aureus",
                                      [CONST_ATBNAME_ASTMRSA,"Vancomycin","Clindamycin","Cefoxitin"],
                                      [CONST_NEWVARNAME_ASTMRSA,"RISVancomycin","RISClindamycin","RISCefoxitin"],
                                      "<i>Staphylococcus aureus</i>"],
    "organism_enterococcus_spp":[20,1 if bisentspp==True else CONST_SPECIFY_CODE_AS_NOTUSEDORG,"Enterococcus spp.",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin"],
                                 "<i>Enterococcus</i> spp."],
    "organism_enterococcus_faecalis":[21,CONST_SPECIFY_CODE_AS_NOTUSEDORG if bisentspp==True else 1,"Enterococcus Faecalis",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin"],
                                 "<i>Enterococcus Faecalis</i>"],
    "organism_enterococcus_faecium":[22,CONST_SPECIFY_CODE_AS_NOTUSEDORG if bisentspp==True else 1,"Enterococcus Faecium",
                                 ["Ampicillin","Vancomycin","Teicoplanin","Linezolid","Daptomycin"],
                                 ["RISAmpicillin","RISVancomycin","RISTeicoplanin","RISLinezolid","RISDaptomycin"],
                                 "<i>Enterococcus Faecium</i>"],
    "organism_streptococcus_pneumoniae":[70,1,"Streptococcus pneumoniae",
                                         ["Penicillin G","Oxacillin","Co-trimoxazole",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Erythromycin","Clindamycin","Levofloxacin"],
                                         ["RISPenicillin_G","RISOxacillin","RISSulfamethoxazole_and_trimethoprim",CONST_NEWVARNAME_AST3GC,"RISCeftriaxone","RISCefotaxime","RISErythromycin","RISClindamycin","RISLevofloxacin"],
                                         "<i>Streptococcus pneumoniae</i>"],
    "organism_salmonella_spp":[80,1,"Salmonella spp.",
                                      [CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Ceftriaxone","Cefotaxime","Ceftazidime",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Ertapenem","Doripenem"],
                                      [CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCeftriaxone","RISCefotaxime","RISCeftazidime",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem"],
                                      "<i>Salmonella</i> spp."],
    "organism_escherichia_coli":[30,1,"Escherichia coli",
                                      ["Gentamicin","Amikacin","Co-trimoxazole","Ampicillin",CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Cefpodoxime","Ceftriaxone","Cefotaxime","Ceftazidime",
                                       "Cefepime",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Ertapenem","Doripenem","Colistin"],
                                      ["RISGentamicin","RISAmikacin","RISSulfamethoxazole_and_trimethoprim","RISAmpicillin",CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCefpodoxime","RISCeftriaxone","RISCefotaxime","RISCeftazidime",
                                       "RISCefepime",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem","RISColistin"],
                                      "<i>Escherichia coli</i>"],
    "organism_klebsiella_pneumoniae":[40,1,"Klebsiella pneumoniae",
                                      ["Gentamicin","Amikacin","Co-trimoxazole",CONST_ATBNAME_ASTFRQ,"Ciprofloxacin","Levofloxacin",CONST_ATBNAME_AST3GC,"Cefpodoxime","Ceftriaxone","Cefotaxime","Ceftazidime",
                                       "Cefepime",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Ertapenem","Doripenem","Colistin"],
                                      ["RISGentamicin","RISAmikacin","RISSulfamethoxazole_and_trimethoprim",CONST_NEWVARNAME_ASTFRQ,"RISCiprofloxacin","RISLevofloxacin",CONST_NEWVARNAME_AST3GC,"RISCefpodoxime","RISCeftriaxone","RISCefotaxime","RISCeftazidime",
                                       "RISCefepime",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISErtapenem","RISDoripenem","RISColistin"],
                                      "<i>Klebsiella pneumoniae</i>"],
    "organism_pseudomonas_aeruginosa":[50,1,"Pseudomonas aeruginosa",
                                      ["Ceftazidime","Ciprofloxacin","Piperacillin/tazobactam",CONST_ATBNAME_ASTAMINOGLY,
                                       "Gentamicin","Amikacin",CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Doripenem","Colistin"],
                                      ["RISCeftazidime","RISCiprofloxacin","RISPiperacillin_and_tazobactam",CONST_NEWVARNAME_ASTAMINOGLY,
                                       "RISGentamicin","RISAmikacin",CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISDoripenem","RISColistin"],
                                      "<i>Pseudomonas aeruginosa</i>"],
    "organism_acinetobacter_baumannii" if bisabom==True else "organism_acinetobacter_spp":[60,1,"Acinetobacter baumannii" if bisabom==True else "Acinetobacter spp.",
                                      ["Tigecycline","Minocycline",CONST_ATBNAME_ASTAMINOGLY,"Gentamicin","Amikacin",
                                       CONST_ATBNAME_ASTCBPN,"Imipenem","Meropenem","Doripenem","Colistin"],
                                      ["RISTigecycline","RISMinocycline",CONST_NEWVARNAME_ASTAMINOGLY,"RISGentamicin","RISAmikacin",
                                       CONST_NEWVARNAME_ASTCBPN,"RISImipenem","RISMeropenem","RISDoripenem","RISColistin"],
                                      "<i>Acinetobacter baumannii</i>" if bisabom==True else "<i>Acinetobacter</i> spp."],
    CONST_ORG_NOGROWTH:[CONST_ORG_NOGROWTH_ORGCAT,0,"No growth",[],[],"No growth"]
    }
"""
def dict_orgwithatb_mortality(bisabom):
    dict_org = get_dict_orgwithatb_mortality(bisabom)
    return dict_org
def get_dict_orgwithatb_mortality(bisabom):
    return {"organism_staphylococcus_aureus":["Staphylococcus aureus",["MRSA","MSSA"],
                                                                [CONST_NEWVARNAME_ASTMRSA,CONST_NEWVARNAME_ASTMRSA],["1","0"]],
                             "organism_enterococcus_spp":["Enterococcus spp.",["Vancomycin-NS","Vancomycin-S"],
                                                                ["NS_Vancomycin","NS_Vancomycin"],["1","0"]],
                             "organism_enterococcus_faecalis":["Enterococcus Faecalis",["Vancomycin-NS","Vancomycin-S"],
                                                                ["NS_Vancomycin","NS_Vancomycin"],["1","0"]],
                             "organism_enterococcus_faecium":["Enterococcus Faecium",["Vancomycin-NS","Vancomycin-S"],
                                                                ["NS_Vancomycin","NS_Vancomycin"],["1","0"]],
                             "organism_streptococcus_pneumoniae":["Streptococcus pneumoniae",["Penicillin-NS","Penicillin-S"],
                                                                ["NS_Penicillin_G","NS_Penicillin_G"],["1","0"]],
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
    dict_org = get_dict_orgwithatb_incidence(bisabom)
    return dict_org
def get_dict_orgwithatb_incidence(bisabom):
    return {"organism_staphylococcus_aureus":["S. aureus",["MRSA"],[CONST_NEWVARNAME_ASTMRSA],["1"],"<i>S. aureus</i>"],
                             "organism_enterococcus_spp":["Enterococcus spp.",["Vancomycin-NS\n<i>Enterococcus</i> spp."],["NS_Vancomycin"],["1"],"<i>Enterococcus</i> spp."],
                             "organism_enterococcus_faecalis":["Enterococcus Faecalis",["Vancomycin-NS\n<i>E. faecalis</i>"],["NS_Vancomycin"],["1"],"<i>E. faecalis</i>"],
                             "organism_enterococcus_faecium":["Enterococcus Faecium",["Vancomycin-NS\n<i>E. faecium</i>"],["NS_Vancomycin"],["1"],"<i>E. faecium</i>"],
                             "organism_streptococcus_pneumoniae":["S. pneumoniae",["Penicillin-NS\n<i>S. pneumoniae</i>"],["NS_Penicillin_G"],["1"],"<i>S. pneumoniae</i>"],
                             "organism_salmonella_spp":["Salmonella spp.",["Fluoroquinolone-NS\n<i>Salmonella</i> spp."],[CONST_NEWVARNAME_ASTFRQ],["1"],"<i>Salmonella</i> spp."],
                             "organism_escherichia_coli":["E. coli",["3GC-NS\n<i>E. coli</i>","Carbapenem-NS\n<i>E. coli</i>"],[CONST_NEWVARNAME_AST3GC,CONST_NEWVARNAME_ASTCBPN],["1","1"],"<i>E. coli</i>"],
                             "organism_klebsiella_pneumoniae":["K. pneumoniae",["3GC-NS\n<i>K. pneumoniae</i>","Carbapenem-NS\n<i>K. pneumoniae</i>"],[CONST_NEWVARNAME_AST3GC,CONST_NEWVARNAME_ASTCBPN],["1","1"],"<i>K. pneumoniae</i>"],
                             "organism_pseudomonas_aeruginosa":["P. aeruginosa",["Carbapenem-NS\n<i>P. aeruginosa</i>"],[CONST_NEWVARNAME_ASTCBPN],["1"],"<i>P. aeruginosa</i>"],
                             "organism_acinetobacter_baumannii" if bisabom==True else "organism_acinetobacter_spp":["A. baumannii" if bisabom==True else "Acinetobacter spp.",
                                                                                                                    ["Carbapenem-NS\n<i>A. baumannii</i>" if bisabom==True else "Carbapenem-NS\n<i>Acinetobacter</i> spp."],
                                                                                                                    [CONST_NEWVARNAME_ASTCBPN],["1"],
                                                                                                                    "<i>A. baumannii</i>" if bisabom==True else "<i>Acinetobacter </i>spp."]
                          } 
# We can convert antibiotic list to read from configuration file that we can configure atb vs organism to set it RIS level that will start consider as AST=1 (Currently I and S -> AST =1)
list_antibiotic = ["Amikacin","Amoxicillin","Amoxicillin_and_clavulanic_acid","Ampicillin","Ampicillin_and_sulbactam","Aztreonam","Cefazolin",
                 "Cefepime","Cefixime","Cefotaxime","Cefotetan","Cefoxitin","Cefpodoxime","Ceftaroline",
                 "Ceftazidime","Ceftriaxone","Cefuroxime","Chloramphenicol","Ciprofloxacin",
                 "Clarithromycin","Clindamycin","Colistin","Dal fopristin_and_quinupristin",
                 "Daptomycin","Doripenem","Doxycycline","Ertapenem","Erythromycin","Fosfomycin",
                 "Fusidic_acid","Gentamicin","Imipenem","Levofloxacin","Linezolid","Meropenem",CONST_ATBNAME_ASTMRSA,"Minocycline",
                 "Moxifloxacin","Nalidixic_acid","Netilmicin","Nitrofurantoin","Oxacillin","Penicillin_G","Piperacillin_and_tazobactam",
                 "Polymyxin_B","Rifampin","Streptomycin","Teicoplanin","Telavancin","Tetracycline","Ticarcillin_and_clavulanic_acid","Tigecycline",
                 "Tobramycin","Trimethoprim","Sulfamethoxazole_and_trimethoprim","Vancomycin","Cefoperazone_and_sulbactam"]
#Annex A
dict_annex_a_listorg = {"organism_burkholderia_pseudomallei":[1,1,"Burkholderia pseudomallei",1,[],"<i>B. pseudomallei</i>"],
                        "organism_brucella_spp":[2,1,"Brucella spp.",1,[],"<i>Brucella </i>spp."],
                        "organism_corynebacterium_diphtheriae":[3,1,"Corynebacterium. diphtheriae",1,[],"<i>C. diphtheriae</i>"],
                        "organism_neisseria_gonorrhoeae":[4,1,"Neisseria. gonorrhoeae",1,[],"<i>N. gonorrhoeae</i>"],
                        "organism_neisseria_meningitidis":[5,1,"Neisseria. meningitidis",1,[],"<i>N. meningitidis</i>"],
                        "organism_non-typhoidal_salmonella_spp":[6,1,"Non-typhoidal Salmonella spp",1,[],"Non-typhoidal <i>Salmonella </i>spp."],
                        "organism_salmonella_paratyphi":[7,1,"Salmonella. Paratyphi",1,[],"<i>S. Paratyphi</i>"],
                        "organism_salmonella_typhi":[8,1,"Salmonella. Typhi",1,[],"<i>S. Typhi</i>"],
                        "organism_shigella_spp":[9,1,"Shigella spp.",1,[],"<i>Shigella </i>spp."],
                        "organism_streptococcus_suis":[10,1,"Streptococcus. suis",1,[],"<i>S. suis</i>"],
                        "organism_vibrio_spp":[11,1,"Vibrio spp.",1,[],"<i>Vibrio </i>spp."],
                        CONST_ORG_NOGROWTH:[CONST_ORG_NOGROWTH_ORGCAT_ANNEX_A,0,"",1,[],""]
                        }
dict_annex_a_spectype = {"specimen_blood":"Blood",
                         "specimen_cerebrospinal_fluid":"CSF",
                         "specimen_genital_swab":"Genital\nswab",
                         "specimen_respiratory_tract":"RTS",
                         "specimen_stool":"Stool",
                         "specimen_urine":"Urine",
                         "specimen_others":"Others"
                        }
#For Report
CONST_MAX_ATBCOUNTFITHALFPAGE = 12
CONST_MAX_ATBCOUNTFITHALFPAGE_MORALITY = 8
CONST_REPORTPAGENUM_MODE = 3 #1 = NORMAL,2 = By section (Need section name), 3 = NORMAL but limit each section maximum page if exceed maximum page it will be 5A, 5B, 5C and so on
list_atbneednote = [CONST_ATBNAME_AST3GC,CONST_ATBNAME_ASTCBPN,CONST_ATBNAME_ASTFRQ,CONST_ATBNAME_ASTAMINOGLY]
dict_atbnote = {CONST_ATBNAME_ASTMRSA:"Methicillin: methicillin, oxacillin, or cefoxitin;",
                CONST_ATBNAME_AST3GC:"3GC=3rd−generation cephalosporin;",
                CONST_ATBNAME_ASTCBPN:"CARBAPENEMS: imipenem, meropenem, ertapenem or doripenem;",
                CONST_ATBNAME_ASTFRQ:"FLUOROQUINOLONES: ciprofloxacin or levofloxacin;",
                CONST_ATBNAME_ASTAMINOGLY:"AMINOGLYCOSIDES: either gentamicin or amikacin;"}
dict_atbnote_sec4_5 =  {"3GC":"3GC=3rd−generation cephalosporin;",
                        "Carbapenem":"CARBAPENEMS: imipenem, meropenem, ertapenem or doripenem;",
                        "Fluoroquinolone":"FLUOROQUINOLONES: ciprofloxacin or levofloxacin;",
                        "AMINOGLYCOSIDE":"AMINOGLYCOSIDES: either gentamicin or amikacin;"}
dict_atbnote_sec6 =  {"3GC":"**3GC-NS [for this section]: NS to any 3rd-generation cephalosporin excluding isolates which are non-susceptible to carbapenem. " + \
                              "***3GC-S [for this section]: S to all 3rd-generation cephalosporin tested excluding isolates which are non-susceptible to carbapenem.;",
                        "Carbapenem":"Carbapenem-NS=NS to any Carbapenem tested;",
                        "Fluoroquinolone":"Fluoroquinolone−NS=NS to any fluoroquinolone tested;"}

