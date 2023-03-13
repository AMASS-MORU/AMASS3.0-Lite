#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate AMR surveillance reports, Supplementary data indicators reports, and Data verification logfile reports systematically.

# Created on 20th April 2022
import pandas as pd #for creating and manipulating dataframe
import matplotlib.pyplot as plt #for creating graph (pyplot)
import matplotlib #for importing graph elements
import numpy as np #for creating arrays and calculations
import seaborn as sns #for creating graph
from pathlib import Path #for retrieving input's path
from datetime import date #for generating today date
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.lib.styles import ParagraphStyle #for setting paragraph style
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT #for setting paragraph style
from reportlab.platypus import * #for plotting graph and tables
from reportlab.lib.colors import * #for importing color palette
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables


# Core pages
pagenumber_intro = ["1","2"]
pagenumber_ava_1 = ["3","4"] #section1; 2 pages
pagenumber_ava_2 = ["5","6","7","8","9","10","11"] #section2; 7 pages

# Available hospital_admission_data.xlsx
pagenumber_ava_3 = ["12","13","14","15","16","17","18","19","20","21","22","23"] #section3; 12 pages
pagenumber_ava_4 = ["24","25","26"] #section4; 3 pages
pagenumber_ava_5 = ["27","28","29","30","31"] #section1; 5 pages
pagenumber_ava_6 = ["32","33","34","35","36","37"] #section6; 6 pages
pagenumber_ava_annexA = ["38","39","40"] #annexA; 3 pages
pagenumber_ava_annexB = ["41","42"] #annexB; 2 pages
pagenumber_ava_other = ["43","44","45","46","47"] #complement; 6 pages

# #AMASS version 1.1
# e_coli       = "<i>Escherichia coli</i>"
# k_pneumoniae = "<i>Klebsiella pneumoniae</i>"
# p_aeruginosa = "<i>Pseudomonas aeruginosa</i>"
# s_aureus     = "<i>Staphylococcus aureus</i>"
# s_pneumoniae = "<i>Streptococcus pneumoniae</i>"
# aci_spp      = "<i>Acinetobacter</i> spp."
# ent_spp      = "<i>Enterococcus</i> spp."
# sal_spp      = "<i>Salmonella</i> spp."

# lst_org_format= [s_aureus, ent_spp, s_pneumoniae, sal_spp, e_coli, k_pneumoniae, p_aeruginosa, aci_spp]
# lst_org_rpt4 = ["<i>S. aureus</i>", 
#                 "<i>Enterococcus</i> spp.", 
#                 "<i>S. pneumoniae</i>", 
#                 "<i>Salmonella</i> spp.", 
#                 "<i>E. coli</i>", 
#                 "<i>K. pneumoniae</i>", 
#                 "<i>P. aeruginosa</i>", 
#                 "<i>Acinetobacter</i> spp."]
# lst_org_rpt4_0 = ["Staphylococcus aureus", 
#                 "Enterococcus spp.", 
#                 "Streptococcus pneumoniae", 
#                 "Salmonella spp.", 
#                 "Escherichia coli", 
#                 "Klebsiella pneumoniae", 
#                 "Pseudomonas aeruginosa", 
#                 "Acinetobacter spp."]
# lst_org_rpt5_0 = ["MRSA", 
#                 "Vancomycin−NS "+lst_org_rpt4_0[1], 
#                 "Penicillin−NS "+lst_org_rpt4_0[2], 
#                 "Fluoroquinolone−NS "+lst_org_rpt4_0[3], 
#                 "3GC−NS "+lst_org_rpt4_0[4], 
#                 "Carbapenem−NS "+lst_org_rpt4_0[4], 
#                 "3GC−NS "+lst_org_rpt4_0[5], 
#                 "Carbapenem−NS "+lst_org_rpt4_0[5], 
#                 "Carbapenem−NS "+lst_org_rpt4_0[6], 
#                 "Carbapenem−NS "+lst_org_rpt4_0[7]]
# lst_org_full  = []
# lst_org_short = []
# for org in lst_org_format:
#     org = org.replace("<i>", "")
#     org = org.replace("</i>", "")
#     org_1 = org.split(" ")                #['Staphylococcus], [aureus']
#     lst_org_full.append(" ".join(org_1)) #['Staphylococcus aureus', ...]

#     if 'spp' in org_1[1]:                #Staphylococcus spp
#         name = org_1[0][0:3]+"_spp"
#     else:                                #Staphylococcus aureus
#         name = org_1[0][0]+"_"+org_1[1]
#     lst_org_short.append(name.lower())  #['s_aureus', ...]

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

def checkpoint(str_filename):
    #Checking file available
    #return : boolean ; True when file is available; False when file is not available
    return Path(str_filename).is_file()

def prepare_section1_table_for_reportlab(df, checkpoint_hosp):
    ##Preparing section1_table for ploting in pdf
    #df: raw section1_table
    #checkpoint_hosp: checkpoint of hospital_admission_data
    #return value: lst of dataframe "section1_table"
    df_sum = df.set_index(df.columns[0])
    df = df.set_index(df.columns[0]).astype(str)
    df.at["Total",df.columns[0]] = round(df_sum.iloc[:,0].sum(skipna=True))
    df = df.reset_index()
    if checkpoint_hosp:
        df.loc[df["Month"]=="Total","Number_of_hospital_records_in_hospital_admission_data_file"] = round(df_sum.iloc[:,1].sum(skipna=True))
    else:
        df["Number_of_hospital_records_in_hospital_admission_data_file"] = ""
        df.loc[df["Month"]=="Total","Number_of_hospital_records_in_hospital_admission_data_file"] = "NA"
    df = df.rename(columns={"Number_of_specimen_in_microbiology_data_file":"Number of specimen\ndata records in\nmicrobiology_data file", "Number_of_hospital_records_in_hospital_admission_data_file":"Number of admission\ndata records in\nhospital_admission_data file"})
    lst_col = [list(df.columns)]
    lst_df = df.values.tolist()
    lst_df = lst_col + lst_df
    return lst_df

def prepare_section2_table_for_reportlab(df_org, df_pat, lst_org, checkpoint_sec2):
    ##Preparing section2_table for ploting in pdf
    #df: raw section2_table
    #checkpoint_hosp: checkpoint of section2_result
    #return value: lst of dataframe "section2_table"
    d_org_core = {"organism_staphylococcus_aureus":"Staphylococcus aureus", 
                "organism_enterococcus_spp":"Enterococcus spp.", 
                "organism_streptococcus_pneumoniae":"Streptococcus pneumoniae", 
                "organism_salmonella_spp":"Salmonella spp.",
                "organism_escherichia_coli":"Escherichia coli", 
                "organism_klebsiella_pneumoniae":"Klebsiella pneumoniae", 
                "organism_pseudomonas_aeruginosa":"Pseudomonas aeruginosa", 
                "organism_acinetobacter_spp":"Acinetobacter spp.", 
                "organism_acinetobacter_baumannii":"Acinetobacter baumannii"}
    df_org = df_org.set_index("Organism").rename(d_org_core)
    df_pat = df_pat.set_index("Organism").rename(d_org_core)
    if checkpoint_sec2:
        #Creating table for page 6
        df_merge = pd.merge(df_org.astype(str), df_pat.astype(str), on="Organism", how="outer").fillna("0").loc[lst_org]
        df_merge_sum = pd.merge(df_org, df_pat, on="Organism", how="outer").fillna(0)
        if "organism_no_growth" in df_merge_sum.index:
            df_merge_sum = df_merge_sum.drop(index=["organism_no_growth"])
        else:
            pass
        #Reformetting Organism name
        d_org_fmt = {}
        style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=10,alignment=TA_LEFT)
        for i in range(len(lst_org)):
            d_org_fmt[lst_org[i]] = Paragraph(prepare_org_core(lst_org[i], text_line=1, text_style="full", text_work="table", text_bold="Y"),style_summary)
        df_merge = df_merge.rename(index=d_org_fmt)
        #Adding Total
        df_merge.loc["Total:","Number_of_blood_specimens_culture_positive_for_the_organism"] = round(df_merge_sum["Number_of_blood_specimens_culture_positive_for_the_organism"].sum())
        df_merge.loc["Total:","Number_of_blood_specimens_culture_positive_deduplicated"]     = round(df_merge_sum["Number_of_blood_specimens_culture_positive_deduplicated"].sum())
        df_merge = df_merge.reset_index().rename(columns={"Number_of_blood_specimens_culture_positive_for_the_organism":"Number of records\nof blood specimens\nculture positive\nfor the organism", 
                                                            "Number_of_blood_specimens_culture_positive_deduplicated":"**Number of patients with\nblood culture positive\nfor the organism\n(de−duplicated)"})
        #Preparing to list
        lst_col = [list(df_merge.columns)]
        lst_df = df_merge.values.tolist()
        lst_df = lst_col + lst_df
    else:
        #Reformetting Organism name
        d_org_fmt = {}
        style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
        for i in range(len(lst_org)):
            d_org_fmt[lst_org[i]] = Paragraph(prepare_org_core(lst_org[i], text_line=1, text_style="full", text_work="table"),style_summary)
        df_merge = pd.DataFrame(index=lst_org + ["Total:"], 
                                columns=["Number of records\nof blood specimens\nculture positive\nfor the organism", 
                                        "**Number of patients with\nblood culture positive\nfor the organism\n(de−duplicated)"])
        lst_df = df_merge.rename(index=d_org_fmt).fillna("NA").values.tolist()
    return lst_df

def prepare_section3_table_for_reportlab(df_pat, lst_org, checkpoint_sec3):
    ##Preparing section3_table for ploting in pdf
    #df: raw section3_table
    #checkpoint_hosp: checkpoint of section3_result
    #return value: lst of dataframe "section3_table"
    lst_org_fmt = []
    style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
    for i in range(len(lst_org)):
        lst_org_fmt.append(Paragraph(prepare_org_core(lst_org[i], text_line=1, text_style="full", text_work="table", text_bold="Y"),style_summary))
    df_pat["Organism"] = lst_org_fmt
    df_pat_sum = df_pat.fillna(0)
    df_pat = df_pat.astype(str)
    df_pat = df_pat.loc[:,["Organism","Number_of_patients_with_blood_culture_positive","Community_origin","Hospital_origin","Unknown_origin"]]
    df_pat.loc["Total","Organism"] = "Total"
    df_pat.loc["Total","Number_of_patients_with_blood_culture_positive"] = str(df_pat_sum["Number_of_patients_with_blood_culture_positive"].sum())
    df_pat.loc["Total","Community_origin"] = str(round(df_pat_sum["Community_origin"].sum()))
    df_pat.loc["Total","Hospital_origin"] = str(round(df_pat_sum["Hospital_origin"].sum()))
    df_pat.loc["Total","Unknown_origin"] = str(round(df_pat_sum["Unknown_origin"].sum()))
    df_pat = df_pat.rename(columns={"Number_of_patients_with_blood_culture_positive":"Number of patients with\nblood culture positive\nfor the organism", 
                                        "Community_origin":"Community\n-origin**","Hospital_origin":"Hospital\n-origin**","Unknown_origin":"Unknown\n-origin***"})
    lst_col = [list(df_pat.columns)]
    lst_df = df_pat.astype(str).values.tolist()
    lst_df = lst_col + lst_df
    return lst_df

def correct_digit(df=pd.DataFrame(),df_col=[]):
    df_new = df.copy().astype(str)
    for idx in df.index:
        for col in df_col:
            if float(df.loc[idx,col]) < 0.05:
                df_new.at[idx,col] = "0"
            elif float(df.loc[idx,col]) >= 0.95:
                df_new.at[idx,col] = str(int(round(df.loc[idx,col])))
            else:
                pass
    return df_new

def create_table_nons(raw_df, org_full, origin="", org_col="Organism", drug_col="Antibiotic",  ns_col="Non-susceptible(N)", total_col="Total(N)"):
    #Selecting rows by organism and parsing table for PDF
    #raw_df : raw dataframe is opened from "Report2_AMR_proportion_table.csv"
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #org_col : column name of organisms
    #drug_col : column name of antibiotics
    #ns_col : column name of number of non-susceptible
    #total_col : column name of total number of patient
    if origin == "": #section2
        sel_df = raw_df.loc[raw_df[org_col]==org_full].set_index(drug_col).fillna(0)
    else: #section3
        sel_df = raw_df.loc[(raw_df[org_col]==org_full) & (raw_df['Infection_origin']==origin)].set_index(drug_col).fillna(0)
    sel_df["%"] = round(sel_df[ns_col]/sel_df[total_col]*100,1).fillna(0)
    sel_df_1 = correct_digit(df=sel_df,df_col=["lower95CI(%)*","upper95CI(%)*","%"])
    # sel_df_1 = sel_df.copy().astype(str)
    # for idx in sel_df.index:
    #     for col in ["lower95CI(%)*","upper95CI(%)*","%"]:
    #         if float(sel_df.loc[idx,col]) < 0.05:
    #             sel_df_1.at[idx,col] = "0"
    #         elif float(sel_df.loc[idx,col]) >= 0.95:
    #             sel_df_1.at[idx,col] = str(int(round(sel_df.loc[idx,col])))
    #         else:
    #             pass
    sel_df_1["% NS (n)"] = sel_df_1["%"].astype(str) + "% (" + sel_df_1[ns_col].astype(str) + "/" + sel_df_1[total_col].astype(str) + ")"
    sel_df_1["95% CI"]  = sel_df_1["lower95CI(%)*"].astype(str) + "% - " + sel_df_1["upper95CI(%)*"].astype(str) + "%"
    sel_df_1 = sel_df_1.loc[:,["% NS (n)", "95% CI"]].reset_index().rename(columns={drug_col:'Antibiotic agent',"% NS (n)":"Proportion of\n NS isolates (n)"})
    sel_df_1 = sel_df_1.replace("0% (0/0)","NA").replace("0% - 0%","-")
    col = pd.DataFrame(list(sel_df_1.columns)).T
    col.columns = list(sel_df_1.columns)
    return col.append(sel_df_1)

def create_graph_nons(raw_df, org_full, org_short, palette, drug_col, perc_col, upper_col, lower_col, origin=""):
    #Creating graph for PDF
    #raw_df : raw dataframe is opened from "Report2_AMR_proportion_table.csv"
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #org_short : short name of organisms for using to retrieve rows by names ex.s_aureus
    #org_col : column name of organisms
    #drug_col : column name of antibiotics
    #perc_col : column name of %Non-susceptible
    #upper_col : column name of upperCI
    #lower_col : column name of lowerCT
    if origin == "": #section2
        sel_df = raw_df.loc[raw_df['Organism']==org_full].set_index(drug_col).fillna(0)
    else: #section3
        sel_df = raw_df.loc[(raw_df['Organism']==org_full) & (raw_df['Infection_origin']==origin)].set_index(drug_col).fillna(0)
        
    if org_full == "Staphylococcus aureus":
        plt.figure(figsize=(6,4))
    if org_full == "Enterococcus spp.":
        plt.figure(figsize=(6,4))
    elif org_full == "Streptococcus pneumoniae" or org_full == "Pseudomonas aeruginosa" or org_full == "Acinetobacter spp." or org_full == "Acinetobacter baumannii":
        plt.figure(figsize=(7,6))
    elif org_full == "Salmonella spp.":
        plt.figure(figsize=(8,7))
    elif org_full == "Escherichia coli" or org_full == "Klebsiella pneumoniae":
        plt.figure(figsize=(8,12))
    
    sns.barplot(data=sel_df.loc[:,perc_col].to_frame().T, palette=palette,orient='h',capsize=.2)
    for drug in sel_df.index:
        ci_lo=sel_df.loc[drug,lower_col]
        ci_hi=sel_df.loc[drug,upper_col]
        if ci_lo == 0 and ci_hi == 0:
            plt.plot([ci_lo,ci_hi],[drug,drug],'|-',color='black',markersize=12,linewidth=0,markeredgewidth=0)
        else:
            plt.plot([ci_lo,ci_hi],[drug,drug],'|-',color='black',markersize=12,linewidth=3,markeredgewidth=3)
    plt.xlim(0, 100)
    plt.xlabel('*Proportion of NS isolates(%)',fontsize=14)
    plt.ylabel('', fontsize=10)
    sns.despine(left=True)
    plt.xticks(fontname='sans-serif',style='normal',fontsize=20)
    plt.yticks(fontname='sans-serif',style='normal',fontsize=20)
    plt.tick_params(top=False, bottom=True, left=False, right=False,labelleft=True, labelbottom=True)
    plt.tight_layout()
    if origin == "":
        plt.savefig('./ResultData/Report2_AMR_' + org_short + '.png', format='png',dpi=180,transparent=True)
    else:
        plt.savefig('./ResultData/Report3_AMR_' + org_short + '_' + origin + '.png', format='png',dpi=180,transparent=True)

def create_num_patient(raw_df, org_full, org_col, ori_col=""):
    #Retrieving number of positive patient for each organism
    #raw_df : raw dataframe is opened from "Report2_page6_patients_under_this_surveillance_by_organism.csv"
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #org_col : column name of organisms
    if ori_col == "": #section2
        temp_table = raw_df.loc[raw_df[org_col]==org_full].values.tolist() #[['Staphylococcus aureus', 100]]
    else: #section3
        temp_table = raw_df.loc[:,['Organism',ori_col]]
        temp_table = temp_table.loc[raw_df[org_col]==org_full].values.tolist() #[['Staphylococcus aureus', 100]]
    return temp_table[0][1] #100

def create_graphpalette(numer_df, numer_col, org_col, org_full, denom_num, cutoff=70.0, origin=""):
    #Function for creating color palette of each organism based on 70.0 cutoff ratio (default)
    #Lower than cutoff : color gainsboro (very light grey)
    #Equal or higher than cutoff : color darkorange
    #numer_df : raw dataframe is opened from "Report2_AMR_proportion_table.csv"
    #numer_col : column name that will be used as numerator
    #org_col : column name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #denom_num : number that will be used as denominator
    #cutoff : default is 70.0% (percentage; 100%)
    if origin == "": #section2
        sel_df = numer_df.loc[numer_df[org_col]==org_full]
    else:
        sel_df = numer_df.loc[(numer_df[org_col]==org_full) & (numer_df['Infection_origin']==origin)]
    palette = []
    for idx in sel_df.index:
        perc = sel_df.loc[idx,numer_col]/denom_num*100
        if perc < cutoff:
            palette.append('gainsboro')
        else:
            palette.append('darkorange')
    return palette

def create_table_surveillance_1(df_raw, lst_org, text_work_drug="N", freq_col="frequency_per_tested", upper_col="frequency_per_tested_uci", lower_col="frequency_per_tested_lci"):
    df_merge = pd.concat([pd.DataFrame(lst_org,columns=["Organism_fmt"]),df_raw],axis=1)
    if text_work_drug == "N":
        df_merge = df_merge.drop(columns=["Organism"]).rename(columns={"Organism_fmt":"Organism"})
    else:
        df_merge = df_merge.drop(columns=["Organism","Priority_pathogen"]).rename(columns={"Organism_fmt":"Organism"})
    
    for c in [freq_col,upper_col,lower_col]:
        df_merge[c+'_1'] = df_merge[c].astype(int)
        for idx in df_merge.index:
            if df_merge.loc[idx,c+'_1'] >= 0.05: #rounding up values of freq, lci, and uci columns
                df_merge.loc[idx,c+'_1'] = df_merge.loc[idx,c+'_1'] + 1
            else:
                df_merge.loc[idx,c+'_1'] = 0
    df_merge["*Frequency (95% CI)"] = df_merge[freq_col+'_1'].astype(str) + \
                                    "\n (" + df_merge[lower_col+'_1'].astype(str) + "-" + \
                                    df_merge[upper_col+'_1'].astype(str) + ")" #creating '*Frequency (95% CI)' columns
    df_merge["*Frequency (95% CI)"] = df_merge["*Frequency (95% CI)"].replace("0\n (0-0)","NA")
    df_merge_1 = df_merge.loc[:,["Organism","*Frequency (95% CI)"]]  #content
    if text_work_drug == "N":
        df_merge_1 = df_merge_1.rename(columns={"Organism":"Pathogens","*Frequency (95% CI)":"*Frequency of infection\n(per 100,000 tested patients;\n95% CI)"})
    else:
        df_merge_1 = df_merge_1.rename(columns={"Organism":"Non-susceptible\n(NS) pathogens","*Frequency (95% CI)":"*Frequency of infection\n(per 100,000 tested patients;\n95% CI)"})
    lst_col = [list(df_merge_1.columns)]
    lst_df = df_merge_1.values.tolist()
    lst_df = lst_col + lst_df
    return lst_df

def create_graph_surveillance_1(df_raw, lst_org, prefix, text_work_drug="N", freq_col="frequency_per_tested", upper_col="frequency_per_tested_uci", lower_col="frequency_per_tested_lci"):
    #Creating graph for PDF
    #raw_df : raw dataframe is opened from "Report2_AMR_proportion_table.csv"
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #upper_col : column name of upperCI
    #lower_col : column name of lowerCT
    if text_work_drug == "Y":
        df_raw = df_raw.set_index('Priority_pathogen')
        palette = ['rebeccapurple','darkorange','firebrick','dodgerblue','saddlebrown','saddlebrown','yellowgreen','yellowgreen','palevioletred','darkkhaki']
    else:
        df_raw = df_raw.set_index('Organism')
        palette = ['rebeccapurple','darkorange','firebrick','dodgerblue','saddlebrown','yellowgreen','palevioletred','darkkhaki']

    plt.figure(figsize=(6.5,12))
    sns.barplot(data=df_raw.loc[:,freq_col].to_frame().T,palette=palette,orient='h',capsize=.2)
    for idx in df_raw.index:
        ci_lo=df_raw.loc[idx,lower_col]
        ci_hi=df_raw.loc[idx,upper_col]
        if ci_lo == 0 and ci_hi == 0:
            plt.plot([ci_lo,ci_hi],[idx,idx],'|-',color='black',markersize=12,linewidth=0,markeredgewidth=0)
        else:
            plt.plot([ci_lo,ci_hi],[idx,idx],'|-',color='black',markersize=12,linewidth=3,markeredgewidth=3)
    plt.locator_params(axis="x", nbins=7) #set number of bins for x-axis
    plt.xlim(0,round(df_raw[upper_col].max())+50)
    plt.xlabel('*Frequency of infection\n(per 100,000 tested patients)',fontsize=14)
    plt.ylabel('', fontsize=10)
    # sns.despine(left=True)
    sns.despine(top=True,right=True) #set REMOVED border line
    plt.xticks(fontname='sans-serif',style='normal',fontsize=14)
    plt.yticks(np.arange(len(lst_org)),lst_org,fontname='sans-serif',style='normal',fontsize=14)
    # plt.yticks("",fontname='sans-serif',style='normal',fontsize=20)
    plt.tight_layout()
    plt.savefig("./ResultData/" + prefix + '.png', format='png',dpi=180,transparent=True)


def prepare_section6_mortality_table_for_reportlab(df_mor):
    ##Preparing section6_mortality_table for reportlab; page 33
    df_mor_com = create_table_perc_mortal_eachorigin(df_mor,'Organism','Infection_origin','Number_of_deaths','Total_number_of_patients','Community-origin')
    df_mor_hos = create_table_perc_mortal_eachorigin(df_mor,'Organism','Infection_origin','Number_of_deaths','Total_number_of_patients','Hospital-origin')
    df_mor_all = pd.concat([df_mor_com,df_mor_hos],axis=1,sort=False).replace(regex=["0% (0/0)","NaN% 0/0"], value="NA")
    df_mor_all.columns = ['Mortality in patients with\nCommunity−origin BSI', 'Mortality in patients with\nHospital−origin BSI']
    return [list(df_mor_all.columns)] + df_mor_all.astype(str).values.tolist()

def create_table_perc_mortal_eachorigin(df,org_col,ori_col,mortal_col,total_col,origin): 
    df = df.loc[df[ori_col]==origin,:].astype({mortal_col:'int32',total_col:'int32'})
    df_amr = df.drop(columns=[ori_col]).groupby([org_col],sort=False).sum()
    df_amr.loc['Total:',:] = df_amr.sum()
    df_amr['perc_mortal'] = round((df_amr[mortal_col]/df_amr[total_col]*100),1).fillna(0)
    df_amr_1 = correct_digit(df=df_amr, df_col=["perc_mortal",mortal_col,total_col])
    # df_amr_1 = df_amr.copy().astype(str)
    # for idx in df_amr.index:
    #     for col in ["perc_mortal",mortal_col,total_col]:
    #         if float(df_amr.loc[idx,col]) < 0.05:
    #             df_amr_1.at[idx,col] = "0"
    #         elif float(df_amr.loc[idx,col]) >= 0.95:
    #             df_amr_1.at[idx,col] = str(int(round(df_amr.loc[idx,col])))
    #         else:
    #             pass
    df_amr_1['perc_mortal_1'] = (df_amr_1['perc_mortal'] + '% (' + df_amr_1[mortal_col].astype(str) + '/' + df_amr_1[total_col].astype(str) + ')').replace('0% (0/0)','NA')
    return df_amr_1.loc[:,"perc_mortal_1"]

def prepare_section6_mortality_table(df_amr, death_col="Number_of_deaths", total_col="Total_number_of_patients", lower_col="Mortality_lower_95ci", upper_col="Mortality_upper_95ci"):
    ##Preparing mortality table; page34-37
    df_amr["Mortality (n)"] = round(df_amr[death_col]/df_amr[total_col]*100,1).fillna(0)
    df_amr_1 = correct_digit(df=df_amr, df_col=["Mortality (n)", lower_col, upper_col])
    # df_amr_1 = df_amr.copy().astype(str)
    # for idx in df_amr.index:
    #     for col in ["Mortality (n)", lower_col, upper_col]:
    #         if float(df_amr.loc[idx,col]) < 0.05:
    #             df_amr_1.at[idx,col] = "0"
    #         elif float(df_amr.loc[idx,col]) >= 0.95:
    #             df_amr_1.at[idx,col] = str(int(round(df_amr.loc[idx,col])))
    #         else:
    #             pass
    df_amr_1["Mortality (n)"] = (df_amr_1["Mortality (n)"]+"% ("+df_amr_1[death_col].astype(str)+"/"+df_amr_1[total_col].astype(str)+")").replace(regex=["0% (0/0)"],value="NA")
    df_amr_1["95% CI"] = (df_amr_1[lower_col].astype(str)+'% - '+df_amr_1[upper_col].astype(str)+"%").replace(regex=["0.0% - 0.0%"],value="-")
    df_amr_1["Antibiotic"] = df_amr_1["Antibiotic"].replace(regex=["3GC-NS"], value="3GC-NS**").replace(regex=["3GC-S"], value="3GC-S***")
    return df_amr_1.loc[:,["Organism","Antibiotic","Infection_origin","Mortality (n)", "95% CI"]].rename(columns={'Antibiotic':'Type of pathogen'})

def create_table_mortal(df, organism, origin, ori_col="Infection_origin", org_col="Organism"):
    df_1 = df.loc[df[ori_col]==origin,:]
    df_2 = df_1.loc[df_1[org_col]==organism].drop(columns=[org_col,ori_col])
    df_col = pd.DataFrame(df_2.columns,index=df_2.columns).T #column name
    df_3 = df_col.append(df_2).replace("0% (0/0)","NA").replace('0% - 0%',"-") #column name + content
    return df_3

def create_graph_mortal_1(df, organism, origin, prefix, org_col="Organism", ori_col="Infection_origin", drug_col="Antibiotic", perc_col="Mortality (n)", lower_col="Mortality_lower_95ci", upper_col="Mortality_upper_95ci"):
    ##Creating graph of mortality
    df_1 = df.loc[df[org_col]==organism,:]
    df_1 = df_1.loc[df[ori_col]==origin,:].replace(regex=["3GC-NS"],value="3GC-NS**").replace(regex=["3GC-S"], value="3GC-S***").set_index(drug_col)
    if organism == "Staphylococcus aureus":
        plt.figure(figsize=(5,2))
    elif organism == "Enterococcus spp.":
        plt.figure(figsize=(6.3,2))
    elif organism == "Streptococcus pneumoniae":
        plt.figure(figsize=(5.6,2))
    elif organism == "Salmonella spp.":
        plt.figure(figsize=(6.5,2))
    elif organism == "Escherichia coli":
        plt.figure(figsize=(6.2,2.7))
    elif organism == "Klebsiella pneumoniae":
        plt.figure(figsize=(6.2,2.7))
    elif organism == "Pseudomonas aeruginosa":
        plt.figure(figsize=(6.2,2))
    elif organism == "Acinetobacter spp." or organism == "Acinetobacter baumannii":
        plt.figure(figsize=(6.2,2))
    sns.barplot(data=df_1.loc[:,perc_col].to_frame().astype(int).T, palette=['darkorange','darkorange'],orient='h',capsize=.2)
    for drug in df_1.index:
        ci_lo=df_1.loc[drug,lower_col]
        ci_hi=df_1.loc[drug,upper_col]
        if ci_lo == 0 and ci_hi == 0:
            plt.plot([ci_lo,ci_hi],[drug,drug],'|-',color='black',markersize=12,linewidth=0,markeredgewidth=0)
        else:
            plt.plot([ci_lo,ci_hi],[drug,drug],'|-',color='black',markersize=12,linewidth=3,markeredgewidth=3)
    plt.xlim(0, 100)
    plt.xlabel('*Mortality (%)',fontsize=16)
    plt.ylabel('', fontsize=10)
    sns.despine(left=True)
    plt.xticks(fontname='sans-serif',style='normal',fontsize=16)
    plt.yticks(fontname='sans-serif',style='normal',fontsize=16)
    plt.tick_params(top=False, bottom=True, left=False, right=False,labelleft=True, labelbottom=True)
    plt.tight_layout()
    plt.savefig("./ResultData/" + prefix + '.png', format='png',dpi=180,transparent=True)



def prepare_section6_numpat_dict(df_mor, origin, origin_col="Infection_origin", total_col="Total_number_of_patients"):
    ##Preparing dataframe of section6_numpat for page 3-6
    df_mor = df_mor.loc[df_mor[origin_col]==origin,["Organism",total_col]] #.drop(columns=["Antibiotic","Infection_origin","Number_of_deaths","Mortality","Mortality_lower_95ci","Mortality_upper_95ci"]).set_index('Organism').astype(int)
    return (df_mor.groupby(['Organism'],sort=False).sum())

def prepare_org_core(str_org, text_line=1, text_style="full", text_work="table", text_work_drug="N", text_bold="N"):
    lst_org = str_org.split(" ")
    ##Preparing capital character for organism
    if len(lst_org) <= 2: #['MRSA'] or ['Brucella', 'spp'] >>> pass
        pass
    else:                 
        lst_org[-2] = lst_org[-2].capitalize()
            
    ##Preparing format of organism
    if text_style == "full":
        pass
    elif text_style == "short":
        if len(lst_org) == 2: 
            if "spp" == lst_org[-1] or "spp." == lst_org[-1]: #['Brucella', 'spp'] >>> pass
                lst_org = [lst_org[0], "spp."]
            else:                                             #['Burkholderia', 'pseudomallei'] >>> ['B.', 'pseudomallei']
                lst_org = [lst_org[0][0] + ".", lst_org[-1]]
        elif len(lst_org) == 3: 
            if "spp" == lst_org[-1] or "spp." == lst_org[-1]: #['Brucella', 'spp'] >>> pass
                pass
            else:                                             #['Burkholderia', 'pseudomallei'] >>> ['B.', 'pseudomallei']
                lst_org = [lst_org[0], lst_org[1][0] + ".", lst_org[-1]]

    ##Preparing organism for applying to work
    for i in range(len(lst_org)):
        if "spp" in lst_org[i] or "spp." in lst_org[i]:
            lst_org[i] = "spp."
        else:
            if text_work == "table":
                if text_work_drug == "N":  #S. pneumoniae (No durg information)
                    lst_org[i] = "<i>" + lst_org[i] + "</i>"    #<i>S.</i> <i>pneumoniae</i>
                else:                   #Penicillin-NS S. pneumoniae
                    if i == 0:          #1st index >>> Penicillin-NS (passed)
                        pass
                    else:               #other index >>> <i>S.</i> <i>pneumoniae</i> (italic added)
                        lst_org[i] = "<i>" + lst_org[i] + "</i>"
            elif text_work == "graph":  
                if text_work_drug == "N":  #S. pneumoniae (No durg information)
                    lst_org[i] = "$" + lst_org[i] + "$"         #$S.$ $pneumoniae$
                else:                   #Penicillin-NS S. pneumoniae
                    if i == 0:          #1st index >>> Penicillin-NS (passed)
                        pass
                    else:               #other index >>> $S.$ $pneumoniae$ (italic added)
                        lst_org[i] = "$" + lst_org[i] + "$"

    ##Preparing organism in line
    if len(lst_org) > 2:
        if text_line == 1: #['Non-typhoidal', 'salmonella', 'spp'] >>> ['Non-typhoidal', 'Salmonella', 'spp']
            lst_org[-2] = lst_org[-2]
        else:             #['Non-typhoidal', 'salmonella', 'spp'] >>> ['Non-typhoidal', '\nSalmonella', 'spp']
            lst_org[-2] = "\n" + lst_org[-2]

    ##Preparing bold organism
    if text_bold=="Y":
        lst_org[0] = "<b>" + lst_org[0]
        lst_org[-1] = lst_org[-1] + "</b>"
    else:
        pass
    return (" ".join(lst_org))

def prepare_org_annexA(str_org, text_line=1, text_style="full", text_work="table"):
    #str_org : Organism with string datatype ex. 'Non-typhoidal salmonella spp'
    #text_line : Number of line for setting organism name format (1--dafault, or 2) ex. 'Non-typhoidal Salmonella spp' or 'Non-typhoidal \nSalmonella spp'
    #text_style : Format type of organism name ("full"--default, or "short") ex. 'Burkholderia pseudomallei or 'B. pseudomallei'
    #text_work : Type of work that organism will be applied ("table"--default, or "graph")
    lst_org = str_org.split(" ")
    ##Preparing capital character for organism
    if len(lst_org) <= 2: #['MRSA'] or ['Brucella', 'spp'] >>> pass
        pass
    else:                 
        lst_org[-2] = lst_org[-2].capitalize()    
    ##Preparing format of organism
    if text_style == "full":
        pass
    elif text_style == "short":
        if len(lst_org) == 2: 
            if "spp" == lst_org[-1] or "spp." == lst_org[-1]: #['Brucella', 'spp'] >>> pass
                lst_org = [lst_org[0], "spp."]
            else:                                             #['Burkholderia', 'pseudomallei'] >>> ['B.', 'pseudomallei']
                lst_org = [lst_org[0][0] + ".", lst_org[-1]] 
    ##Preparing organism for applying to work
    for i in range(len(lst_org)):
        if "spp" in lst_org[i] or "spp." in lst_org[i]:
            lst_org[i] = "spp."
        else:
            if text_work == "table":
                lst_org[i] = "<i>" + lst_org[i] + "</i>"
            elif text_work == "graph":
                lst_org[i] = "$" + lst_org[i] + "$"    
    ##Preparing organism in line
    if len(lst_org) > 2:
        if text_line == 1: #['Non-typhoidal', 'salmonella', 'spp'] >>> ['Non-typhoidal', 'Salmonella', 'spp']
            lst_org[-2] = lst_org[-2]
        else:             #['Non-typhoidal', 'salmonella', 'spp'] >>> ['Non-typhoidal', '\nSalmonella', 'spp']
            lst_org[-2] = "\n" + lst_org[-2]
    return (" ".join(lst_org))

def prepare_annexA_numpat_table_for_reportlab(df_pat, lst_org):
    annexA_org_page2 = pd.DataFrame(lst_org,columns=["Organism_fmt"])
    df_pat = df_pat.rename(columns={"Total_number_of_patients":"Total number\nof patients*", 
                                    "Number_of_patients_with_blood_positive_deduplicated":"Blood",
                                    "Number_of_patients_with_csf_positive_deduplicated":"CSF",
                                    "Number_of_patients_with_genitLal_swab_positive_deduplicated":"Genital\nswab", 
                                    "Number_of_patients_with_rts_positive_deduplicated":"RTS", 
                                    "Number_of_patients_with_stool_positive_deduplicated":"Stool", 
                                    "Number_of_patients_with_urine_positive_deduplicated":"Urine", 
                                    "Number_of_patients_with_others_positive_deduplicated":"Others"})
    df_pat_2 = pd.concat([annexA_org_page2,df_pat],axis=1).drop(columns=["Organism"]).rename(columns={"Organism_fmt":"Pathogens"})
    df_pat_2.iloc[-1,0] = "Total"
    return [list(df_pat_2.columns)] + df_pat_2.values.tolist()

def prepare_annexA_mortality_table_for_reportlab(df_mor, lst_org, death_col="Number_of_deaths", total_col="Total_number_of_patients", lower_col="Mortality_lower_95ci", upper_col="Mortality_upper_95ci"):
    df_mor["Mortality(%)"] = round(df_mor[death_col]/df_mor[total_col]*100,1).fillna(0)
    # df_mor_1 = df_mor.copy().astype(str)
    df_mor_1 = correct_digit(df=df_mor,df_col=["Mortality(%)",upper_col,lower_col])
    df_mor_1[["Mortality (n)","95% CI"]] = ""
    # for idx in df_mor.index:
    #     mortal = ""
    #     upper = ""
    #     lower = ""
        # if float(df_mor_1.loc[idx,"Mortality(%)"]) < 0.05:
        #     mortal = "0"
        # elif float(df_mor_1.loc[idx,"Mortality(%)"]) >= 0.95:
        #     mortal = str(int(round(df_mor.loc[idx,"Mortality(%)"])))
        # else:
        #     mortal = str(round(df_mor.loc[idx,"Mortality(%)"],1))
            
        # if float(df_mor_1.loc[idx, upper_col]) < 0.05:
        #     upper = "0"
        # elif float(df_mor_1.loc[idx,upper_col]) >= 0.95:
        #     upper = str(int(round(df_mor.loc[idx,upper_col])))
        # else:
        #     upper = str(round(df_mor.loc[idx,upper_col],1))

        # if float(df_mor_1.loc[idx,lower_col]) < 0.05:
        #     lower = "0"
        # elif float(df_mor_1.loc[idx,lower_col]) >= 0.95:
        #     lower = str(int(round(df_mor.loc[idx,lower_col])))
        # else:
        #     lower = str(round(df_mor.loc[idx,lower_col],1))
        # df_mor_1.at[idx,"Mortality (n)"] = str(mortal)+"% ("+str(df_mor_1.loc[idx,death_col])+"/"+str(df_mor_1.loc[idx,total_col])+")"
        # df_mor_1.at[idx,'95% CI'] = str(lower)+'% - '+str(upper)+'%'
    df_mor_1["Mortality (n)"] = (df_mor_1["Mortality(%)"]+"% ("+df_mor_1[death_col].astype(str)+"/"+df_mor_1[total_col].astype(str)+")").replace(regex=["0% (0/0)"],value="NA")
    df_mor_1['95% CI'] = (df_mor_1[lower_col].astype(str)+'% - '+df_mor_1[upper_col].astype(str)+"%").replace(regex=["0.0% - 0.0%"],value="-")
    df_mor_1 = df_mor_1.loc[:,["Organism","Mortality (n)","95% CI"]].replace("0% (0/0)","NA").replace("0% - 0%","-")
    df_mor_2 = pd.concat([lst_org,df_mor_1],axis=1).drop(columns=["Organism"]).rename(columns={"Organism_fmt":"Pathogens"}) #Adding orgaism_fmt to table
    return [list(df_mor_2.columns)] + df_mor_2.values.tolist()

import numpy as np
def create_annexA_mortality_graph(df_mor, lst_org, death_col="Number_of_deaths", total_col="Total_number_of_patients", lower_col="Mortality_lower_95ci", upper_col="Mortality_upper_95ci"):
    df_mor["Mortality(%)"] = round(df_mor[death_col]/df_mor[total_col]*100,1).fillna(0)
    df_mor = df_mor.set_index('Organism')
    palette = ['darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange']
    plt.figure(figsize=(5.0,10))
    sns.barplot(data=df_mor.loc[:,'Mortality(%)'].astype(float).to_frame().T, palette=palette,orient='h',capsize=.2)
    for org in df_mor.index:
        ci_lo=df_mor.loc[org,lower_col]
        ci_hi=df_mor.loc[org,upper_col]
        if ci_lo == 0 and ci_hi == 0:
            plt.plot([ci_lo,ci_hi],[org,org],'|-',color='black',markersize=12,linewidth=0,markeredgewidth=0)
        else:
            plt.plot([ci_lo,ci_hi],[org,org],'|-',color='black',markersize=12,linewidth=3,markeredgewidth=3)
    plt.xlim(0, 100)
    plt.ylabel('', fontsize=10)
    sns.despine(top=True, right=True)
    plt.xticks(fontname='sans-serif',style='normal',fontsize=14)
    plt.yticks(np.arange(len(lst_org)),lst_org,fontname='sans-serif',style='normal',fontsize=14)
    plt.tick_params(top=False, bottom=True, left=True, right=False,labelleft=True, labelbottom=True)
    plt.tight_layout()
    plt.savefig('./ResultData/AnnexA_mortality.png', format='png',dpi=300,transparent=True)

def prepare_annexB_summary_table_for_reportlab(df,indi_col="Indicators",total_col="Total(%)",cri_col="Critical_priority(%)",high_col="High_priority(%)",med_col="Medium_priority(%)"):
    df = df.fillna("NA").loc[:,[indi_col,total_col, cri_col,high_col,med_col]]
    df.at[:,indi_col] = ["Blood culture\ncontamination rate*", "Proportion of notifiable\nantibiotic-pathogen\ncombinations**","Proportion of isolates with\ninfrequent phenotypes or\npotential errors in AST results\n***"]
    for idx in df.index:
        for col in [1,2,3,4]:
            df.iloc[idx,col] = df.iloc[idx,col].replace("(","\n(")
    df = df.rename(columns={total_col:"Total\n(n)",cri_col:"Critical priority\n(n)",high_col:"High priority\n(n)",med_col:"Medium priority\n(n)"})
    df_col = [[indi_col,"Number of observations","","",""],["","Total\n(n)","Critical priority\n(n)","High priority\n(n)","Medium priority\n(n)"]]
    return df_col + df.values.tolist()

def prepare_annexB_summary_table_bymonth_for_reportlab(df, col_month="month", 
                                                       col_rule1="summary_blood_culture_contamination_rate(%)", 
                                                       col_rule2="summary_proportion_of_notifiable_antibiotic-pathogen_combinations(%)", 
                                                       col_rule3="summary_proportion_of_potential_errors_in_the_AST_results(%)"):
    df = df.fillna("NA").loc[:,[col_month, col_rule1, col_rule2, col_rule3]].rename(columns={col_month:"Month", 
                                                                                            col_rule1:"Blood culture\ncontamination rate\n(n)*",
                                                                                            col_rule2:"Proportion of notifiable\nantibiotic-pathogen combinations\n(n)**", 
                                                                                            col_rule3:"Proportion of isolates with\ninfrequent phenotypes or\npotential errors in AST results\n(n)***"})
    return [list(df.columns)] + df.values.tolist()

def check_number_org_annexA(lst_org):
    if len(lst_org) <= 12:
        num_left = 12-len(lst_org)
        i = 0
        while i < num_left:
            lst_org.append("")
            i += 1 
    else:
        lst_org = lst_org[:12]
    return lst_org

#Assigning Not available of NA
#Return value: assigned variables
def assign_na_toinfo(str_info, coverpage=False):
    if str_info == "empty001_micro" or str_info == "empty001_hosp" or str_info == "NA" or str_info == "" or str_info != str_info:
        if coverpage:
            str_info = "Not available"
        else:
            str_info = "NA"
    else:
        str_info = str(str_info)
    return str_info

def report_title(c,title_name,pos_x,pos_y,font_color,font_size=20):
    c.setFont("Helvetica-Bold", font_size) # define a large bold Helvetica
    c.setFillColor(font_color) #define font color
    c.drawString(pos_x,pos_y,title_name)

def report_context(c,context_list,pos_x,pos_y,wide,height,font_size=10,font_align=TA_JUSTIFY,line_space=18,left_indent=0):
    context_list_style = []
    style = ParagraphStyle('normal',fontName='Helvetica',leading=line_space,fontSize=font_size,leftIndent=left_indent,alignment=font_align)
    for cont in context_list:
        cont_1 = Paragraph(cont, style)
        context_list_style.append(cont_1)
    f = Frame(pos_x,pos_y,wide,height,showBoundary=0)
    return f.addFromList(context_list_style,c)

def report_todaypage(c,pos_x,pos_y,footer_information):
    c.setFont("Helvetica", 9) # define a large bold font
    c.setFillColor('#3e4444')
    c.drawString(pos_x,pos_y,footer_information)

def report1_table(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(1,1),(-1,-1),'Helvetica-BoldOblique'),
                           ('FONTSIZE',(0,0),(-1,-1),11),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.darkblue),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('ALIGN',(0,0),(-3,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report2_table(df):
    return Table(df,style=[('FONT',(0,0),(-1,0),'Helvetica-Bold'), 
                           ('FONT',(1,1),(-1,-1),'Helvetica-BoldOblique'),
                           ('FONT',(0,-1),(0,-1),'Helvetica-Bold'), 
                           ('FONTSIZE',(0,0),(-1,-1),11),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.darkblue),
                           ('ALIGN',(0,0),(-1,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report3_table(df):
    return Table(df,style=[('FONT',(0,0),(-1,0),'Helvetica-Bold'), 
                           ('FONT',(0,1),(-1,-1),'Helvetica-BoldOblique'),
                           ('FONTSIZE',(0,0),(-1,-1),11),
                           ('TEXTCOLOR',(0,1),(-1,-1),colors.darkblue),
                           ('ALIGN',(0,0),(-1,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report2_table_nons(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report_table_annexA_page1(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),14),
                           ('GRID',(0,0),(-1,-1),0.5,colors.white),
                           ('ALIGN',(0,0),(-1,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], 
                 colWidths=[3.0*inch,3.0*inch])

def report_table_annexA_page2(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE'), 
                           ('ALIGN',(0,-1),(0,-1),'LEFT')], 
                 colWidths=[1.3*inch,0.9*inch,0.5*inch,0.5*inch,0.5*inch,0.5*inch,0.5*inch,0.5*inch,0.5*inch])

def report_table_annexA_page3(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], 
                 colWidths=[1.2*inch,1.0*inch,1.0*inch])

def report_table_annexB(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report_table_annexB_page1(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica'),
                           ('FONT',(0,0),(-1,1),'Helvetica-Bold'),
                           ('FONTSIZE',(0,0),(-1,-1),10),
                           ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                           ('ALIGN',(0,2),(0,-1),'LEFT'),
                           ('ALIGN',(0,0),(0,1),'CENTER'),
                           ('ALIGN',(1,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE'), 
                           ('SPAN',(0,0),(0,1)), 
                           ('SPAN',(1,0),(-1,0))])

