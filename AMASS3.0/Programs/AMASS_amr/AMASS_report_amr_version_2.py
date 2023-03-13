#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate AMR surveillance reports, Supplementary data indicators reports, and Data verification logfile reports systematically.

# Created on 20th April 2022
import logging #for creating error_log
import re #for manipulating regular expression
import pandas as pd #for creating and manipulating dataframe
import matplotlib.pyplot as plt #for creating graph (pyplot)
import matplotlib #for importing graph elements
from pathlib import Path #for retrieving input's path
from PIL import Image #for importing image
import seaborn as sns #for creating graph
from datetime import date #for generating today date
from reportlab.lib.pagesizes import A4 #for setting PDF size
from reportlab.pdfgen import canvas #for creating PDF page
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.lib.styles import ParagraphStyle #for setting paragraph style
from reportlab.lib.enums import TA_LEFT, TA_CENTER #for setting paragraph style
from reportlab.platypus import * #for plotting graph and tables
from reportlab.lib.colors import * #for importing color palette
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
from AMASS_report_amr_function_version_2 import * #for importing amr functions
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables


# Create a logging instance
logger = logging.getLogger('AMASS_report_amr_version_2.py')
logger.setLevel(logging.INFO)
# Assign a file-handler to that instance
fh = logging.FileHandler("./error_report_amr.txt")
fh.setLevel(logging.ERROR)
# Format your logs (optional)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
# Add the handler to your logging instance
logger.addHandler(fh)

##paragraph variable
iden1_op = "<para leftindent=\"35\">"
iden2_op = "<para leftindent=\"70\">"
iden3_op = "<para leftindent=\"105\">"
iden_ed = "</para>"
bold_blue_ital_op = "<b><i><font color=\"#000080\">"
bold_blue_ital_ed = "</font></i></b>"
bold_blue_op = "<b><font color=\"#000080\">"
bold_blue_ed = "</font></b>"
green_op = "<font color=darkgreen>"
green_ed = "</font>"
add_blankline = "<br/>"
tab1st = "&nbsp;"
tab4th = "&nbsp;&nbsp;&nbsp;&nbsp;"

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

#AMASS version 1.1
e_coli       = "<i>Escherichia coli</i>"
k_pneumoniae = "<i>Klebsiella pneumoniae</i>"
p_aeruginosa = "<i>Pseudomonas aeruginosa</i>"
s_aureus     = "<i>Staphylococcus aureus</i>"
s_pneumoniae = "<i>Streptococcus pneumoniae</i>"
aci_spp      = "<i>Acinetobacter</i> spp."
ent_spp      = "<i>Enterococcus</i> spp."
sal_spp      = "<i>Salmonella</i> spp."

path_result = "./ResultData/"
path_input = "./"
sec1_res_i = "Report1_page3_results.csv"
sec1_num_i = "Report1_page4_counts_by_month.csv"
sec2_res_i = "Report2_page5_results.csv"
sec2_amr_i = "Report2_AMR_proportion_table.csv"
sec2_org_i = "Report2_page6_counts_by_organism.csv"
sec2_pat_i = "Report2_page6_patients_under_this_surveillance_by_organism.csv"
sec3_res_i = "Report3_page12_results.csv"
sec3_amr_i = "Report3_table.csv"
sec3_pat_i = "Report3_page13_counts_by_origin.csv"
sec4_res_i = "Report4_page24_results.csv"
sec4_blo_i = "Report4_frequency_blood_samples.csv"
sec4_pri_i = "Report4_frequency_priority_pathogen.csv"
sec5_res_i = "Report5_page27_results.csv"
sec5_com_i = "Report5_incidence_blood_samples_community_origin.csv"
sec5_hos_i = "Report5_incidence_blood_samples_hospital_origin.csv"
sec5_com_amr_i = "Report5_incidence_blood_samples_community_origin_antibiotic.csv"
sec5_hos_amr_i = "Report5_incidence_blood_samples_hospital_origin_antibiotic.csv"
sec6_res_i = "Report6_page32_results.csv"
sec6_mor_byorg_i = "Report6_mortality_byorganism.csv"
sec6_mor_i = "Report6_mortality_table.csv"
secA_res_i = "AnnexA_page39_results.csv"
secA_pat_i = "AnnexA_patients_with_positive_specimens.csv"
secA_mor_i = "AnnexA_mortlity_table.csv"
secB_blo_i = "AnnexB_proportion_table_blood.csv"
secB_blo_mon_i = "AnnexB_proportion_table_blood_bymonth.csv"

##### dictionary_for_microbiology_data ##### 01/09/22
dict_ = pd.DataFrame()
try:
    dict_ = pd.read_excel(path_input + "dictionary_for_microbiology_data.xlsx").iloc[:,:2].fillna("")
except:
    try:
        dict_ = pd.read_csv(path_input + "dictionary_for_microbiology_data.csv").iloc[:,:2].fillna("")
    except:
        dict_ = pd.read_csv(path_input + "dictionary_for_microbiology_data.csv", encoding="windows-1252").iloc[:,:2].fillna("")
dict_.columns = ["amass_name","user_name"]

try:
    check_abaumannii = True if "organism_acinetobacter_baumannii" in list(dict_.loc[dict_["amass_name"]=="acinetobacter_spp_or_baumannii","user_name"]) else False
except:
    check_abaumannii = False
lst_org_rpt4_0 = []
if check_abaumannii:
    aci_spp      = "<i>Acinetobacter baumannii</i>"
    lst_org_rpt4_0 = ["Staphylococcus aureus", 
                      "Enterococcus spp.", 
                      "Streptococcus pneumoniae", 
                      "Salmonella spp.", 
                      "Escherichia coli", 
                      "Klebsiella pneumoniae", 
                      "Pseudomonas aeruginosa", 
                      "Acinetobacter baumannii"]
else:
    aci_spp        = "<i>Acinetobacter</i> spp."
    lst_org_rpt4_0 = ["Staphylococcus aureus", 
                      "Enterococcus spp.", 
                      "Streptococcus pneumoniae", 
                      "Salmonella spp.", 
                      "Escherichia coli", 
                      "Klebsiella pneumoniae", 
                      "Pseudomonas aeruginosa", 
                      "Acinetobacter spp."]
lst_org_format= [s_aureus, ent_spp, s_pneumoniae, sal_spp, e_coli, k_pneumoniae, p_aeruginosa, aci_spp]
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
lst_org_rpt5_0 = ["MRSA", 
                "Vancomycin−NS "+lst_org_rpt4_0[1], 
                "Penicillin−NS "+lst_org_rpt4_0[2], 
                "Fluoroquinolone−NS "+lst_org_rpt4_0[3], 
                "3GC−NS "+lst_org_rpt4_0[4], 
                "Carbapenem−NS "+lst_org_rpt4_0[4], 
                "3GC−NS "+lst_org_rpt4_0[5], 
                "Carbapenem−NS "+lst_org_rpt4_0[5], 
                "Carbapenem−NS "+lst_org_rpt4_0[6], 
                "Carbapenem−NS "+lst_org_rpt4_0[7]]
lst_org_full  = []
lst_org_short = []
for org in lst_org_format:
    org = org.replace("<i>", "")
    org = org.replace("</i>", "")
    org_1 = org.split(" ")                #['Staphylococcus], [aureus']
    lst_org_full.append(" ".join(org_1)) #['Staphylococcus aureus', ...]

    if 'spp' in org_1[1]:                #Staphylococcus spp
        name = org_1[0][0:3]+"_spp"
    else:                                #Staphylococcus aureus
        name = org_1[0][0]+"_"+org_1[1]
    lst_org_short.append(name.lower())  #['s_aureus', ...]

today = date.today()
checkpoint_mic = checkpoint(path_input + "microbiology_data.xlsx") or checkpoint(path_input + "microbiology_data.csv")
checkpoint_hos = checkpoint(path_input + "hospital_admission_data.xlsx") or checkpoint(path_input + "hospital_admission_data.csv")
##### Configuration #####
try:
    config = pd.read_excel(path_input + "Configuration/Configuration.xlsx")
except:
    try:
        config = pd.read_csv(path_input + "Configuration/Configuration.csv")
    except:
        config = pd.read_csv(path_input + "Configuration/Configuration.csv", encoding="windows-1252")
##### section1 #####
if check_config(config, "amr_surveillance_section1"):
    print ("AMR surveillance report - checkpoint section1")
    try:
        sec1_res = pd.read_csv(path_result + sec1_res_i).fillna("NA")
        sec1_T = pd.read_csv(path_result + sec1_num_i)
        sec1_T = prepare_section1_table_for_reportlab(sec1_T,checkpoint_hos)
    except Exception as e:
        logger.exception(e)
        pass
else:
    pass
##### section2 #####
if check_config(config, "amr_surveillance_section2") and checkpoint(path_result + sec2_res_i):
    print ("AMR surveillance report - checkpoint section2")
    sec2_T_org1 = []
    sec2_T_org2 = []
    sec2_T_org3 = []
    sec2_T_org4 = []
    sec2_T_org5 = []
    sec2_T_org6 = []
    sec2_T_org7 = []
    sec2_T_org8 = []
    try:
        sec2_res = pd.read_csv(path_result + sec2_res_i).fillna("NA")
        sec2_T_amr = pd.read_csv(path_result + sec2_amr_i)
        sec2_T_org = pd.read_csv(path_result + sec2_org_i)
        sec2_T_pat = pd.read_csv(path_result + sec2_pat_i)
        ##Section 2; Page 1
        sec2_merge = prepare_section2_table_for_reportlab(sec2_T_org, sec2_T_pat, lst_org_full, checkpoint(path_result + sec2_res_i))
        ##Section 2; Page 2-6
        #Retriving numper of positive patient of each organism
        #Creating AMR grapgh of each organism
        #Creating AMR table of each organism
        lst_numpat = []
        for i in range(len(lst_org_full)):
            numpat = create_num_patient(sec2_T_pat, lst_org_full[i], 'Organism')
            lst_numpat.append(numpat)
            palette = create_graphpalette(sec2_T_amr,'Total(N)','Organism',lst_org_full[i],numpat,cutoff=70.0)
            sec2_G = create_graph_nons(sec2_T_amr,lst_org_full[i], lst_org_short[i],palette,'Antibiotic','Non-susceptible(%)','upper95CI(%)*','lower95CI(%)*')
        sec2_org1 = create_table_nons(sec2_T_amr,lst_org_full[0]).values.tolist()
        sec2_org2 = create_table_nons(sec2_T_amr,lst_org_full[1]).values.tolist()
        sec2_org3 = create_table_nons(sec2_T_amr,lst_org_full[2]).values.tolist()
        sec2_org4 = create_table_nons(sec2_T_amr,lst_org_full[3]).values.tolist()
        sec2_org5 = create_table_nons(sec2_T_amr,lst_org_full[4]).values.tolist()
        sec2_org6 = create_table_nons(sec2_T_amr,lst_org_full[5]).values.tolist()
        sec2_org7 = create_table_nons(sec2_T_amr,lst_org_full[6]).values.tolist()
        sec2_org8 = create_table_nons(sec2_T_amr,lst_org_full[7]).values.tolist()
    except Exception as e:
        logger.exception(e)
        pass
else:
    pass
##### section3 #####
if check_config(config, "amr_surveillance_section3") and checkpoint(path_result + sec3_res_i):
    print ("AMR surveillance report - checkpoint section3")
    try:
        sec3_res = pd.read_csv(path_result + sec3_res_i).fillna("NA")
        sec3_amr = pd.read_csv(path_result + sec3_amr_i)
        sec3_pat = pd.read_csv(path_result + sec3_pat_i)
        sec3_pat_1 = sec3_pat.copy()
        ##Section 3; Page 1
        sec3_pat = pd.read_csv(path_result + sec3_pat_i).drop(columns=["Number_of_patients_with_blood_culture_positive_merged_with_hospital_data_file"])
        sec3_pat_sum = sec3_pat.copy().fillna(0)
        sec3_pat = sec3_pat.astype(str)
        sec3_pat_1 = sec3_pat.astype(str)
        sec3_pat.loc["Total","Organism"] = "Total"
        sec3_pat.loc["Total","Number_of_patients_with_blood_culture_positive"] = str(sec3_pat_sum["Number_of_patients_with_blood_culture_positive"].sum())
        sec3_pat.loc["Total","Community_origin"] = str(round(sec3_pat_sum["Community_origin"].sum()))
        sec3_pat.loc["Total","Hospital_origin"] = str(round(sec3_pat_sum["Hospital_origin"].sum()))
        sec3_pat.loc["Total","Unknown_origin"] = str(round(sec3_pat_sum["Unknown_origin"].sum()))
        sec3_pat = sec3_pat.rename(columns={"Number_of_patients_with_blood_culture_positive":"Number of patients with\nblood culture positive\nfor the organism", 
                                            "Community_origin":"Community\n-origin**","Hospital_origin":"Hospital\n-origin**","Unknown_origin":"Unknown\n-origin***"})
        sec3_col = pd.DataFrame(list(sec3_pat.columns),index=sec3_pat.columns).T
        sec3_pat_val = sec3_col.append(sec3_pat).drop(columns=["Organism"]).values.tolist()
        ##Section 3; Page 3-12
        sec3_lst_numpat = []    
        for i in range(len(lst_org_full)):
            for ori in ['Community_origin','Hospital_origin']:
                numpat = create_num_patient(sec3_pat_1.astype(str), lst_org_full[i], 'Organism', ori)
                sec3_lst_numpat.append(numpat)
                palette = create_graphpalette(sec3_amr,'Total(N)','Organism',lst_org_full[i],float(numpat),cutoff=70.0,origin=ori[:-7])
                sec3_G = create_graph_nons(sec3_amr,lst_org_full[i], lst_org_short[i],palette,'Antibiotic','Non-susceptible(%)','upper95CI(%)*','lower95CI(%)*', origin=ori[:-7])
        sec3_org1_com = create_table_nons(sec3_amr,lst_org_full[0],"Community").values.tolist()
        sec3_org1_hos = create_table_nons(sec3_amr,lst_org_full[0],"Hospital").values.tolist()
        sec3_org2_com = create_table_nons(sec3_amr,lst_org_full[1],"Community").values.tolist()
        sec3_org2_hos = create_table_nons(sec3_amr,lst_org_full[1],"Hospital").values.tolist()
        sec3_org3_com = create_table_nons(sec3_amr,lst_org_full[2],"Community").values.tolist()
        sec3_org3_hos = create_table_nons(sec3_amr,lst_org_full[2],"Hospital").values.tolist()
        sec3_org4_com = create_table_nons(sec3_amr,lst_org_full[3],"Community").values.tolist()
        sec3_org4_hos = create_table_nons(sec3_amr,lst_org_full[3],"Hospital").values.tolist()
        sec3_org5_com = create_table_nons(sec3_amr,lst_org_full[4],"Community").values.tolist()
        sec3_org5_hos = create_table_nons(sec3_amr,lst_org_full[4],"Hospital").values.tolist()
        sec3_org6_com = create_table_nons(sec3_amr,lst_org_full[5],"Community").values.tolist()
        sec3_org6_hos = create_table_nons(sec3_amr,lst_org_full[5],"Hospital").values.tolist()
        sec3_org7_com = create_table_nons(sec3_amr,lst_org_full[6],"Community").values.tolist()
        sec3_org7_hos = create_table_nons(sec3_amr,lst_org_full[6],"Hospital").values.tolist()
        sec3_org8_com = create_table_nons(sec3_amr,lst_org_full[7],"Community").values.tolist()
        sec3_org8_hos = create_table_nons(sec3_amr,lst_org_full[7],"Hospital").values.tolist()
    except Exception as e:
        logger.exception(e)
        pass
else:
    pass
##### section4 #####
if check_config(config, "amr_surveillance_section4") and checkpoint(path_result + sec4_res_i):
    print ("AMR surveillance report - checkpoint section4")
    try:
        sec4_res = pd.read_csv(path_result + sec4_res_i).fillna("NA")
        style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
        ##Section 4; Page 2
        lst_org_rpt4_graph = []
        lst_org_rpt4_table = []
        for i in range(len(lst_org_rpt4_0)): #page2
            lst_org_rpt4_graph.append(prepare_org_core(lst_org_rpt4_0[i], text_line=1, text_style="short", text_work="graph"))
            lst_org_rpt4_table.append(Paragraph(prepare_org_core(lst_org_rpt4_0[i], text_line=1, text_style="short", text_work="table"),style_summary))
        sec4_blo = pd.read_csv(path_result + sec4_blo_i)
        sec4_blo_1 = create_table_surveillance_1(sec4_blo, lst_org_rpt4_table)
        create_graph_surveillance_1(sec4_blo, lst_org_rpt4_graph, 'Report4_frequency_blood')
        ##Section 4; Page 3
        lst_org_rpt4_graph_page3 = []
        lst_org_rpt4_table_page3 = []
        for i in range(len(lst_org_rpt5_0)): #page3
            lst_org_rpt4_graph_page3.append(prepare_org_core(lst_org_rpt5_0[i], text_line=2, text_style="short", text_work="graph", text_work_drug="Y"))
            lst_org_rpt4_table_page3.append(Paragraph(prepare_org_core(lst_org_rpt5_0[i], text_line=2, text_style="short", text_work="table", text_work_drug="Y").replace(" \n<i>","<br/><i>"),style_summary))
        sec4_pat = pd.read_csv(path_result + sec4_pri_i)
        sec4_pat_1 = create_table_surveillance_1(sec4_pat, lst_org_rpt4_table_page3, text_work_drug="Y")
        create_graph_surveillance_1(sec4_pat, lst_org_rpt4_graph_page3, 'Report4_frequency_pathogen', text_work_drug="Y")
    except Exception as e:
        logger.exception(e)
        print("--Section 4 error---")
        print(e)
        pass
else:
    pass
##### section5 #####
if check_config(config, "amr_surveillance_section5") and checkpoint(path_result + sec5_res_i):
    print ("AMR surveillance report - checkpoint section5")
    try:
        sec5_res = pd.read_csv(path_result + sec5_res_i).fillna("NA")
        style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
        lst_org_rpt4_graph = []
        lst_org_rpt4_table = []
        for i in range(len(lst_org_rpt4_0)): #page2
            lst_org_rpt4_graph.append(prepare_org_core(lst_org_rpt4_0[i], text_line=1, text_style="short", text_work="graph"))
            lst_org_rpt4_table.append(Paragraph(prepare_org_core(lst_org_rpt4_0[i], text_line=1, text_style="short", text_work="table"),style_summary))
        lst_org_rpt4_graph_page3 = []
        lst_org_rpt4_table_page3 = []
        for i in range(len(lst_org_rpt5_0)): #page3
            lst_org_rpt4_graph_page3.append(prepare_org_core(lst_org_rpt5_0[i], text_line=2, text_style="short", text_work="graph", text_work_drug="Y"))
            lst_org_rpt4_table_page3.append(Paragraph(prepare_org_core(lst_org_rpt5_0[i], text_line=2, text_style="short", text_work="table", text_work_drug="Y").replace(" \n<i>","<br/><i>"),style_summary))
        ##Section 5; Page 2
        sec5_com = pd.read_csv(path_result + sec5_com_i)
        sec5_com_1 = create_table_surveillance_1(sec5_com, lst_org_rpt4_table)
        create_graph_surveillance_1(sec5_com, lst_org_rpt4_graph, 'Report5_incidence_community')
        ##Section 5; Page 3
        sec5_hos = pd.read_csv(path_result + sec5_hos_i)
        sec5_hos_1 = create_table_surveillance_1(sec5_hos, lst_org_rpt4_table)
        create_graph_surveillance_1(sec5_hos, lst_org_rpt4_graph, 'Report5_incidence_hospital')
        ##Section 5; Page 4
        sec5_com_amr = pd.read_csv(path_result + sec5_com_amr_i)
        sec5_com_amr_1 = create_table_surveillance_1(sec5_com_amr, lst_org_rpt4_table_page3)
        create_graph_surveillance_1(sec5_com_amr, lst_org_rpt4_graph_page3, 'Report5_incidence_community_antibiotic', text_work_drug="Y")
        ##Section 5; Page 5
        sec5_hos_amr = pd.read_csv(path_result + sec5_hos_amr_i)
        sec5_hos_amr_1 = create_table_surveillance_1(sec5_hos_amr, lst_org_rpt4_table_page3)
        create_graph_surveillance_1(sec5_hos_amr, lst_org_rpt4_graph_page3, 'Report5_incidence_hospital_antibiotic', text_work_drug="Y")
    except Exception as e:
        logger.exception(e)
        pass
else:
    pass
##### section6 #####
if check_config(config, "amr_surveillance_section6") and checkpoint(path_result + sec6_res_i):
    print ("AMR surveillance report - checkpoint section6")
    try:
        sec6_res = pd.read_csv(path_result + sec6_res_i).fillna("NA")
        sec6_mor = pd.read_csv(path_result + sec6_mor_i)
        sec6_mor_byorg = pd.read_csv(path_result + sec6_mor_byorg_i).fillna(0)
        ##Creating table for page33
        sec6_mor_all = prepare_section6_mortality_table_for_reportlab(sec6_mor_byorg)
        #### section6; page2-5 #####
        sec6_mor_1 = prepare_section6_mortality_table(sec6_mor)
        sec6_mor_com_1 = create_table_mortal(sec6_mor_1, lst_org_full[0], "Community-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_hos_1 = create_table_mortal(sec6_mor_1, lst_org_full[0], "Hospital-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_com_2 = create_table_mortal(sec6_mor_1, lst_org_full[1], "Community-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_hos_2 = create_table_mortal(sec6_mor_1, lst_org_full[1], "Hospital-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_com_3 = create_table_mortal(sec6_mor_1, lst_org_full[2], "Community-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_hos_3 = create_table_mortal(sec6_mor_1, lst_org_full[2], "Hospital-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_com_4 = create_table_mortal(sec6_mor_1, lst_org_full[3], "Community-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_hos_4 = create_table_mortal(sec6_mor_1, lst_org_full[3], "Hospital-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_com_5 = create_table_mortal(sec6_mor_1, lst_org_full[4], "Community-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_hos_5 = create_table_mortal(sec6_mor_1, lst_org_full[4], "Hospital-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_com_6 = create_table_mortal(sec6_mor_1, lst_org_full[5], "Community-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_hos_6 = create_table_mortal(sec6_mor_1, lst_org_full[5], "Hospital-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_com_7 = create_table_mortal(sec6_mor_1, lst_org_full[6], "Community-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_hos_7 = create_table_mortal(sec6_mor_1, lst_org_full[6], "Hospital-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_com_8 = create_table_mortal(sec6_mor_1, lst_org_full[7], "Community-origin").replace(regex=["\r"],value="").values.tolist()
        sec6_mor_hos_8 = create_table_mortal(sec6_mor_1, lst_org_full[7], "Hospital-origin").replace(regex=["\r"],value="").values.tolist()
        #Creating graph for page 34-37
        sec6_G = sec6_mor.copy().replace(regex=["\r"],value="")
        sec6_G["Mortality (n)"] = round(sec6_G["Number_of_deaths"]/sec6_G["Total_number_of_patients"]*100,1).fillna(0)
        for i in range(len(lst_org_full)):
            for ori in ['Community-origin','Hospital-origin']:
                create_graph_mortal_1(sec6_G,lst_org_full[i],ori,'Report6_mortality_'+lst_org_short[i]+'_'+ori[:-7].lower())
        #number of patient
        sec6_numpat_com = prepare_section6_numpat_dict(sec6_mor_byorg, "Community-origin")
        sec6_numpat_hos = prepare_section6_numpat_dict(sec6_mor_byorg, "Hospital-origin")
    except Exception as e:
        logger.exception(e)
        pass
else:
    pass  
##### AnnexA #####
if check_config(config, "amr_surveillance_annexA"):
    print ("AMR surveillance report - checkpoint annex A")
    if checkpoint(path_result + secA_res_i):
        try:
            secA_res = pd.read_csv(path_result + secA_res_i).fillna("NA")
            secA_pat = pd.read_csv(path_result + secA_pat_i).fillna("NA")
            #Preparing organism for AnnexA page1-3
            style_normal = ParagraphStyle('normal',fontName='Helvetica',fontSize=11,alignment=TA_LEFT)
            style_small = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
            annexA_org = secA_pat["Organism"].tolist()
            annexA_org.remove("Total")
            lst_page1 = []
            lst_page2 = []
            for i in range(len(annexA_org)):
                lst_page1.append(Paragraph("- " + prepare_org_annexA(annexA_org[i], text_line=1, text_style="full", text_work="table").replace("<i>Non-typhoidal</i>","Non-typhoidal").replace("<i>paratyphi</i>","Paratyphi").replace("<i>typhi</i>","Typhi"), style_normal))
                lst_page2.append(Paragraph(prepare_org_annexA(annexA_org[i], text_line=2, text_style="short", text_work="table").replace("<i>Non-typhoidal</i>","Non-typhoidal").replace("<i>paratyphi</i>","Paratyphi").replace("<i>typhi</i>","Typhi"), style_small))
            #Creating formatted organism for AnnexA page1
            annexA_org_page1 = pd.DataFrame()
            lst_page1 = check_number_org_annexA(lst_page1)
            annexA_org_page1["1-6"] = lst_page1[:6]
            annexA_org_page1["7-12"] = lst_page1[6:]
            annexA_org_page1 = annexA_org_page1.values.tolist()
            #Creating formatted organism for AnnexB's table on page2
            secA_pat2 = prepare_annexA_numpat_table_for_reportlab(secA_pat, lst_page2)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        pass
    if checkpoint(path_result + secA_mor_i):
        try:
            ##Creating table for page40 (Parsing)
            secA_mortal = pd.read_csv(path_result + secA_mor_i).fillna(0)
            #Preparing organism for AnnexA page3
            style = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
            annexA_org = secA_mortal["Organism"].tolist()
            lst_page3_table = []
            lst_page3_graph = []
            for i in range(len(annexA_org)):
                lst_page3_table.append(Paragraph(prepare_org_annexA(annexA_org[i], text_line=2, text_style="short", text_work="table").replace("<i>Non-typhoidal</i>","Non-typhoidal").replace("<i>paratyphi</i>","Paratyphi").replace("<i>typhi</i>","Typhi"),style))
                lst_page3_graph.append(prepare_org_annexA(annexA_org[i], text_line=2, text_style="short", text_work="graph").replace("$Non-typhoidal$","Non-typhoidal").replace("$paratyphi$","Paratyphi").replace("$typhi$","Typhi"))
            #Creating formatted organism for AnnexB's table on page3
            annexA_org_page3 = pd.DataFrame(lst_page3_table,columns=["Organism_fmt"])
            ##Creating table for AnnexB page3
            secA_mortal3 = prepare_annexA_mortality_table_for_reportlab(secA_mortal, annexA_org_page3)
            ##Creating graph for AnnexB page3
            create_annexA_mortality_graph(secA_mortal,lst_page3_graph)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        pass
##### AnnexB #####
if check_config(config, "amr_surveillance_annexB"):
    print ("AMR surveillance report - checkpoint annex B")
    secB_blo_1 = []
    secB_blo_bymonth = []
    ##Creating table for AnnexB page1 (Parsing)
    if checkpoint(path_result + secB_blo_i):
        try:
            secB_blo = pd.read_csv(path_result + secB_blo_i) #.fillna("NA")
            secB_blo_1 = prepare_annexB_summary_table_for_reportlab(secB_blo)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        pass
    ##Creating table for AnnexB page2 (Parsing)
    if checkpoint(path_result + secB_blo_mon_i):
        try:
            secB_blo_bymonth = pd.read_csv(path_result + secB_blo_mon_i) #.fillna("NA")
            secB_blo_bymonth = prepare_annexB_summary_table_bymonth_for_reportlab(secB_blo_bymonth)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        pass

def cover(section1_result=pd.DataFrame(),today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    add_blankline = "<br/>"
    ##variables
    if len(section1_result) > 0:
        hospital_name   = assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Hospital_name","Values"].tolist()[0]), coverpage=True)
        country_name    = assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Country","Values"].tolist()[0]), coverpage=True)
        contact_person  = assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Contact_person","Values"].tolist()[0]), coverpage=True)
        contact_address = assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Contact_address","Values"].tolist()[0]), coverpage=True)
        contact_email   = assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Contact_email","Values"].tolist()[0]), coverpage=True)
        notes   = str(section1_result.loc[section1_result["Parameters"]=="notes_on_the_cover","Values"].tolist()[0])
        spc_date_start  = assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0]), coverpage=True)
        spc_date_end    = assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0]), coverpage=True)
    else:
        hospital_name    = "NA"
        country_name    = "NA"
        contact_person  = "NA"
        contact_address = "NA"
        contact_email   = "NA"
        notes           = "NA"
        spc_date_start  = "NA"
        spc_date_end    = "NA"

    ##text
    cover_1_1 = "<b>Hospital name:</b>  " + bold_blue_op + hospital_name + bold_blue_ed
    cover_1_2 = "<b>Country name:</b>  " + bold_blue_op + country_name + bold_blue_ed
    cover_1_3 = "<b>Data from:</b>"
    cover_1_4 = bold_blue_op + str(spc_date_start) + " to " + str(spc_date_end) + bold_blue_ed
    cover_1 = [cover_1_1,cover_1_2,add_blankline+cover_1_3, cover_1_4]
    cover_2_1 = "<b>Contact person:</b>  " + bold_blue_op + contact_person + bold_blue_ed
    cover_2_2 = "<b>Contact address:</b>  " + bold_blue_op + contact_address + bold_blue_ed
    cover_2_3 = "<b>Contact email:</b>  " + bold_blue_op + contact_email + bold_blue_ed
    cover_2_4 = "<b>Generated on:</b>  " + bold_blue_op + today + bold_blue_ed
    cover_2 = [cover_2_1,cover_2_2,cover_2_3,cover_2_4]
    if notes == "" or notes == "NA":
        pass
    else:
        cover_2 = cover_2 + ["<b>Notes:</b>  " + notes]
    ########### COVER PAGE ############
    c.setFillColor('#FCBB42')
    c.rect(0,590,800,20, fill=True, stroke=False)
    c.setFillColor(forestgreen)
    c.rect(0,420,800,150, fill=True, stroke=False)
    report_title(c,'Antimicrobial Resistance (AMR)',0.7*inch, 515,'white',font_size=28)
    report_title(c,'Surveillance report',0.7*inch, 455,'white',font_size=28)
    report_context(c,cover_1, 0.7*inch, 3.0*inch, 460, 180, font_size=18,line_space=26)
    report_context(c,cover_2, 0.7*inch, 0.5*inch, 460, 120, font_size=11)
    c.showPage()

def generatedby(section1_result=pd.DataFrame()):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden_ed = "</para>"
    add_blankline = "<br/>"
    ##variables
    if len(section1_result) > 0:
        hospital_name   = assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Hospital_name","Values"].tolist()[0]), coverpage=False)
        country_name    = assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Country","Values"].tolist()[0]), coverpage=False)
        spc_date_start  = assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0]), coverpage=False)
        spc_date_end    = assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0]), coverpage=False)
    else:
        hospital_name    = "NA"
        country_name    = "NA"
        spc_date_start  = "NA"
        spc_date_end    = "NA"

    generatedby_1_1  = "Generated by"
    generatedby_1_2  = "AutoMated tool for Antimicrobial resistance Surveillance System (AMASS) version 2.0"
    generatedby_1_3  = "(released on 16 May 2022)"
    generatedby_1_5  = "The AMASS application is available under the Creative Commons Attribution 4.0 International Public License (CC BY 4.0). The application can be downloaded at : <u><link href=\"https://www.amass.website\" color=\"blue\"fontName=\"Helvetica\">https://www.amass.website</link></u>"
    generatedby_1_6  = "The AMASS application used microbiology_data and hospital_admission_data files that are stored in the same folder as the application (AMASS.bat) to generate this report."
    generatedby_1_7  = "The goal of the AMASS application is to enable hospitals with microbiology data available in electronic formats to analyze their own data and generate AMR surveillance reports promptly. If hospital admission date data are available, the reports will additionally be stratified by infection origin (community−origin or hospital−origin). If mortality data (such as patient discharge outcome data) are available, a report on mortality involving AMR infection will be added."
    generatedby_1_8  = "This automatically generated report has limitations, and requires users to understand those limitations and use the summary data in the report with careful interpretation."
    generatedby_1_9  = "A valid report could have local implications and much wider benefits if shared with national and international organizations."
    generatedby_1_10  = "This automatically generated report is under the jurisdiction of the hospital to copy, redistribute, and share with any individual or organization."
    generatedby_1_11 = "This automatically generated report contains no patient identifier, similar to standard reports on cumulative antimicrobial susceptibility."
    generatedby_1_12 = "For any query on AMASS, please contact:"
    generatedby_1_13 = "Chalida Rangsiwutisak (chalida@tropmedres.ac),"
    generatedby_1_14 = "Cherry Lim (cherry@tropmedres.ac), and"
    generatedby_1_15 = "Direk Limmathurotsakul (direk@tropmedres.ac)"
    generatedby_1 = ["<b>" + generatedby_1_1 + "</b>", 
                    generatedby_1_2, 
                    generatedby_1_3,  
                    iden1_op + add_blankline + generatedby_1_5 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_6 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_7 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_8 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_9 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_10 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_11 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_12 + iden_ed, 
                    iden1_op + generatedby_1_13 + iden_ed, 
                    iden1_op + generatedby_1_14 + iden_ed, 
                    iden1_op + generatedby_1_15 + iden_ed]
    generatedby_2_1 = "Suggested title for citation:"
    generatedby_2_2 = "Antimicrobial resistance surveillance report, " + hospital_name + ","
    generatedby_2_3 = country_name + ", "+ str(spc_date_start) + " to " + str(spc_date_end) + "."
    generatedby_2 = ["<b>" + generatedby_2_1 + "</b>", 
                    generatedby_2_2, 
                    generatedby_2_3]
    ########### GENERATED BY ##########
    report_context(c,generatedby_1, 1.0*inch, 0.7*inch, 460, 700, font_size=11, line_space=16)
    report_context(c,generatedby_2, 1.0*inch, 0.6*inch, 460, 80, font_size=11, line_space=16)
    c.showPage()

def tableofcontent():
    content_0 = 'Introduction'
    content_1 = 'Section [1]: Data overview'
    content_2 = 'Section [2]: Isolate−based surveillance report'
    content_3 = 'Section [3]: Isolate−based surveillance report with stratification by infection origin'
    content_4 = 'Section [4]: Sample−based surveillance report'
    content_5 = 'Section [5]: Sample−based surveillance report with stratification by infection origin'
    content_6 = 'Section [6]: Mortality involving AMR and antimicrobial−susceptible infections'
    content_7 = 'Annex A: Supplementary report on notifiable bacterial infections'
    content_8 = 'Annex B: Supplementary report on data indicators'
    content_9 = 'Methods'
    content_10 = 'Acknowledgements'
    content = [content_0, content_1, content_2, content_3, content_4, content_5, content_6, content_7, content_8, content_9, content_10]
    content_page = ['01', '03', '05', '12', '24', '27', '32', '38', '41', '43', '48']
    ############# CONTENT #############
    report_title(c,'Content',1.07*inch, 10.5*inch,'#3e4444', font_size=16)
    report_context(c,content, 1.0*inch, 6.0*inch, 435, 300, font_size=11)
    report_context(c,content_page, 7.0*inch, 6.0*inch, 30, 300, font_size=11)
    c.showPage()

def introduction(lst_pagenumber=pagenumber_intro, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden_ed = "</para>"
    add_blankline = "<br/>"
    ##Page1
    intro_page1_1 = "Antimicrobial resistance (AMR) is a global health crisis [1]. "+ \
                "The report by Lord Jim O'Neill estimated that 700,000 global deaths could be attributable to AMR in 2015, and projected that the annual death toll could reach 10 million by 2050 [1]. "+ \
                "However, data of AMR surveillance from low and middle−income countries (LMICs) are scarce [1,2], and data of mortality associated with AMR infections are rarely available. "+ \
                "A recent study estimated that 19,000 deaths are attributable to AMR infections in Thailand annually, using routinely available microbiological and hospital databases [3]. "+ \
                "The study also proposed that hospitals in LMICs should utilize routinely available microbiological and hospital admission databases to generate reports on AMR surveillance systematically [3]."
    intro_page1_2_1 = "Reports on AMR surveillance can have a wide range of benefits [2]; including"
    intro_page1_2_2 = "− characterization of the frequency of resistance and organisms in different facilities and regions;"
    intro_page1_2_3 = "− prospective and retrospective information on emerging public health threats;"
    intro_page1_2_4 = "− evaluation and optimization of local and national standard treatment guidelines;"
    intro_page1_2_5 = "− evaluation of the impact of interventions beyond antimicrobial guidelines that aim to reduce AMR; and"
    intro_page1_2_6 = "− data sharing with national and international organizations to support decisions on resource allocation for interventions against AMR and to inform the implementation of action plans at national and global levels."
    intro_page1_3 = "When reporting AMR surveillance results, it is generally recommended that (a) duplicate results of bacterial isolates are removed, and (b) reports are stratified by infection origin (community−origin or hospital−origin), if possible [2]. "+ \
                    "Many hospitals in LMICs lack time and resources needed to analyze the data (particularly to deduplicate data and to generate tables and figures), write the reports, and to release the data or reports [4]."
    intro_page1_4 = "AutoMated tool for Antimicrobial resistance Surveillance System (AMASS) was developed as an offline, open−access and easy−to−use application that allows a hospital to perform data analysis independently and generate isolate−based and sample−based surveillance reports stratified by infection origin from routinely collected electronic databases. The application was built in R, which is a free software environment. The application has been placed within a user−friendly interface that only requires the user to double−click on the application icon. The AMASS application can be downloaded at: <u><link href=\"https://www.amass.website\" color=\"blue\"fontName=\"Helvetica\">https://www.amass.website</link></u>"
    intro_page1 = [intro_page1_1, 
                add_blankline + intro_page1_2_1, 
                iden1_op + intro_page1_2_2 + iden_ed, 
                iden1_op + intro_page1_2_3 + iden_ed, 
                iden1_op + intro_page1_2_4 + iden_ed, 
                iden1_op + intro_page1_2_5 + iden_ed, 
                iden1_op + intro_page1_2_6 + iden_ed, 
                add_blankline + intro_page1_3, 
                add_blankline + intro_page1_4]
    ##Page2
    intro_page2_1_1 = "The AMASS version 2.0 additionally generates reports on notifiable bacterial diseases in Annex A and on data indicators (including proportion of contaminants and discordant AST results) in Annex B for the \"microbiology_data\" file that is used to generate this report. A careful review of the Annex B could help readers and data owners to identify potential errors in the microbiology data used to generate the report."
    intro_page2_1_2 = "The AMASS version 2.0 also separately generates Supplementary data indictors report (in PDF and Excel formats) in a new folder “Report_with_patient_identifiers” to support users to check and validate records with notifiable bacteria, notifiable antibiotic-pathogen combinations, infrequent phenotypes or potential errors in the AST results at the local level. The identifiers listed include hospital number and specimen collection date. The files are generated in a separate folder “Report_with_patient_identifiers” so that it is clear that users should not share or transfer the Supplementary Data Indictors report (in PDF and Excel format) to any party outside of the hospital without data security management and confidential agreement."
    intro_page2_1 = [intro_page2_1_1, add_blankline + intro_page2_1_2]
    intro_page2_2_1 = "References:"
    intro_page2_2_2 = "[1] O'Neill J. (2014) Antimicrobial resistance: tackling a crisis for the health and wealth of nations. Review on antimicrobial resistance. http://amr−review.org. (accessed on 3 Dec 2018)."
    intro_page2_2_3 = "[2] World Health Organization (2018) Global Antimicrobial Resistance Surveillance System (GLASS) Report. Early implantation 2016−2017. http://apps.who.int/iris/bitstream/handle/10665/259744/9789241513449−eng.pdf. (accessed on 3 Dec 2018)"
    intro_page2_2_4 = "[3] Lim C., et al. (2016) Epidemiology and burden of multidrug−resistant bacterial infection in a developing country. Elife 5: e18082."
    intro_page2_2_5 = "[4] Ashley EA, Shetty N, Patel J, et al. Harnessing alternative sources of antimicrobial resistance data to support surveillance in low−resource settings. J Antimicrob Chemother. 2019; 74(3):541−546."
    intro_page2_2_6 = "[5] Clinical and Laboratory Standards Institute (CLSI). Analysis and Presentation of Cumulative Antimicrobial Susceptibility Test Data, 4th Edition. 2014. (accessed on 21 Jan 2020)"
    intro_page2_2_7 = "[6] European Antimicrobial Resistance Surveillance Network (EARS−Net). Antimicrobial resistance (AMR) reporting protocol 2018. (accessed on 21 Jan 2020)"
    intro_page2_2_8 = "[7] European Committee on Antimicrobial Susceptibility Testing (EUCAST). www.eucast.org (accessed on 21 Jan 2020)"
    intro_page2_2 = ["<b>" + intro_page2_2_1 + "</b>", 
                    intro_page2_2_2, 
                    intro_page2_2_3, 
                    intro_page2_2_4, 
                    intro_page2_2_5, 
                    intro_page2_2_6, 
                    intro_page2_2_7, 
                    intro_page2_2_8]
    ########## INTRO: PAGE1 ###########
    report_title(c,'Introduction',1.07*inch, 10.5*inch,'#3e4444', font_size=16)
    report_context(c,intro_page1, 1.07*inch, 1.0*inch, 460, 650, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()
    ########## INTRO: PAGE2 ###########
    report_context(c,intro_page2_1, 1.07*inch, 4.5*inch, 460, 450, font_size=11)
    report_context(c,intro_page2_2, 1.07*inch, 0.5*inch, 460, 270, font_size=9,font_align=TA_LEFT,line_space=14) #Reference
    u = inch/10.0
    c.setLineWidth(2)
    c.setStrokeColor(black)
    p = c.beginPath()
    p.moveTo(70,315) # start point (x,y)
    p.lineTo(7.2*inch,315) # end point (x,y)
    c.drawPath(p, stroke=1, fill=1)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.rotate(90)
    c.showPage()

def section1(section1_result, section1_table, lst_pagenumber=pagenumber_ava_1, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start  = str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0])
    spc_date_end    = str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0])
    blo_num         = str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Number_of_records"),"Values"].tolist()[0])
    if checkpoint(path_input+"hospital_admission_data.xlsx") or checkpoint(path_input+"hospital_admission_data.csv"):
        hos_date_start = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0])
        hos_date_end   = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0])
        hos_num        = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Number_of_records"),"Values"].tolist()[0])
        patient_days   = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Patient_days"),"Values"].tolist()[0])
        patient_days_his = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Patient_days_his"),"Values"].tolist()[0])
    else:
        hos_date_start = "NA"
        hos_date_end   = "NA"
        hos_num        = "NA"
        patient_days   = "NA"
        patient_days_his = "NA"

    ##Page1
    section1_page1_1_1 = "An overview of the data detected by the AMASS application is generated by default. "+ \
                    "The summary is based on the raw data files saved within the same folder as the application file (AMASS.bat)."
    section1_page1_1_2 = "Please review and validate this section carefully before proceeds to the next section."
    section1_page1_1 = [section1_page1_1_1, 
                        add_blankline + section1_page1_1_2]
    section1_page1_2_1 = "The microbiology_data file (stored in the same folder as the application file) had:"
    section1_page1_2_2 = bold_blue_ital_op + blo_num + bold_blue_ital_ed + " specimen data records with collection dates ranging from "
    section1_page1_2_3 = bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + " to " + bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section1_page1_2_4 = "The hospital_admission_data file (stored in the same folder as the application file) had:"
    section1_page1_2_5 = bold_blue_ital_op + hos_num + bold_blue_ital_ed + " admission data records with hospital admission dates ranging from "
    section1_page1_2_6 = bold_blue_ital_op + hos_date_start + bold_blue_ital_ed + " to " + bold_blue_ital_op + hos_date_end + bold_blue_ital_ed
    section1_page1_2_7 = "The total number of patient-days was " + bold_blue_ital_op + patient_days + bold_blue_ital_ed + "."
    section1_page1_2_8 = "The total number of patient-days at risk of BSI of hospital-origin was " + bold_blue_ital_op + patient_days_his + bold_blue_ital_ed + "."
    section1_page1_2 = [section1_page1_2_1, 
                        iden1_op + "<i>" + section1_page1_2_2 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section1_page1_2_3 + "</i>" + iden_ed, 
                        add_blankline + section1_page1_2_4, 
                        iden1_op + "<i>" + section1_page1_2_5 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section1_page1_2_6 + "</i>" + iden_ed,
                        iden1_op + add_blankline + section1_page1_2_7 + iden_ed,
                        iden1_op + add_blankline + section1_page1_2_8 + iden_ed]
    section1_page1_3_1 = "[1] If the periods of the data in microbiology_data and hospital_admission_data files are not similar, " + \
                        "the automatically−generated report should be interpreted with caution. " + \
                        "The AMASS generates the reports based on the available data."
    section1_page1_3_2 = "[2] A patient is defined as at risk of BSI of hospital-origin when the patient is admitted to the hospital for more than two calendar days with calendar day one equal to the day of admission."
    section1_page1_3 = [green_op + section1_page1_3_1 + green_ed, 
                        green_op + section1_page1_3_2 + green_ed]
    ##Page2
    section1_page2_1_1 = "Data was stratified by month to assist detection of missing data, and verification of whether the month distribution of data records in microbiology_data file and hospital_admission_data file reflected the microbiology culture frequency and admission rate of the hospital, respectively. " + \
                        "For example if the number of specimens in the microbiology_data file reported below is lower than what is expected, please check the raw data file and data dictionary files."
    section1_page2_1 = [section1_page2_1_1]
    section1_page2_2_1 = "[1] Additional general demographic data will be made available in the next version of the AMASS application."
    section1_page2_2 = [green_op + section1_page2_2_1 + green_ed]
    ######### SECTION1: PAGE1 #########
    report_title(c,'Section [1]: Data overview',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,section1_page1_1, 1.0*inch, 7.3*inch, 460, 150, font_size=11)
    report_title(c,'Results',1.07*inch, 7.5*inch,'#3e4444',font_size=12)
    report_context(c,section1_page1_2, 1.0*inch, 4.0*inch, 460, 250, font_size=11)
    report_title(c,'Note:',1.07*inch, 3.5*inch,'darkgreen',font_size=12)
    report_context(c,section1_page1_3, 1.0*inch, 1.8*inch, 460, 120, font_size=11)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of "+ lastpage)
    c.showPage()
    ######### SECTION1: PAGE2 #########
    report_title(c,'Reporting period by months:',1.07*inch, 10.5*inch,'#3e4444',font_size=12)
    report_context(c,section1_page2_1, 1.0*inch, 8.0*inch, 460, 150, font_size=11)
    table_draw = report1_table(section1_table)
    table_draw.wrapOn(c, 600, 400)
    table_draw.drawOn(c, 1.5*inch, 4.3*inch)
    report_title(c,'Note:',1.07*inch, 3.0*inch,'darkgreen',font_size=12)
    report_context(c,section1_page2_2, 1.0*inch, 2.3*inch, 460, 50, font_size=11)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.showPage()

def section2(result_table, summary_table, org1, org2, org3, org4, org5, org6, org7, org8, 
             lst_org, lst_pagenumber=pagenumber_ava_2, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variable
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden3_op = "<para leftindent=\"105\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start  = str(result_table.loc[result_table["Parameters"]=="Minimum_date","Values"].tolist()[0])
    spc_date_end    = str(result_table.loc[result_table["Parameters"]=="Maximum_date","Values"].tolist()[0])
    blo_num         = str(result_table.loc[result_table["Parameters"]=="Number_of_blood_specimens_collected","Values"].tolist()[0])
    blo_num_neg     = str(result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_negative","Values"].tolist()[0])
    blo_num_pos     = str(result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_positive","Values"].tolist()[0])
    blo_num_pos_org = str(result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_positive_for_organism_under_this_survey","Values"].tolist()[0])
    ##Page1
    section2_page1_1_1 = "An isolate−based surveillance report is generated by default, even if the hospital_admission_data file is unavailable. "+ \
                        "This is to enable hospitals with only microbiology data available to utilize the de−duplication and report generation functions of AMASS. " + \
                        "This report is without stratification by origin of infection."
    section2_page1_1_2 = "The report generated by the AMASS application version 2.0 includes only blood samples. " + \
                        "The next version of AMASS will include other specimen types, including cerebrospinal fluid (CSF), urine, stool, and other specimens."
    section2_page1_1 = [section2_page1_1_1, 
                        add_blankline + section2_page1_1_2]
    section2_page1_2 = [iden1_op + "− " + lst_org[0] + iden_ed, 
                        iden1_op + "− " + lst_org[1] + iden_ed, 
                        iden1_op + "− " + lst_org[2] + iden_ed, 
                        iden1_op + "− " + lst_org[3] + iden_ed, 
                        iden1_op + "− " + lst_org[4] + iden_ed, 
                        iden1_op + "− " + lst_org[5] + iden_ed, 
                        iden1_op + "− " + lst_org[6] + iden_ed, 
                        iden1_op + "− " + lst_org[7] + iden_ed]
    section2_page1_3_1 = "The microbiology_data file had:"
    section2_page1_3_2 = "Sample collection dates ranged from " + \
                        bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + " to " + bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section2_page1_3_3 = "Number of records of blood specimens collected within the above date range:"
    section2_page1_3_4 = blo_num + " blood specimens records"
    section2_page1_3_5 = "Number of records of blood specimens with *negative culture (no growth):"
    section2_page1_3_6 = blo_num_neg + " blood specimens records"
    section2_page1_3_7 = "Number of records of blood specimens with culture positive for a microorganism:"
    section2_page1_3_8 = blo_num_pos + " blood specimens records"
    section2_page1_3_9 = "Number of records of blood specimens with culture positive for organism under this survey:"
    section2_page1_3_10 = blo_num_pos_org + " blood specimens records"
    section2_page1_3 = [section2_page1_3_1, 
                        iden1_op + "<i>"             + section2_page1_3_2 + "</i>" + iden_ed, 
                        iden1_op + "<i>"             + section2_page1_3_3 + "</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section2_page1_3_4 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>"             + section2_page1_3_5 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section2_page1_3_6 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>"             + section2_page1_3_7 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section2_page1_3_8 + bold_blue_ital_ed + iden_ed, 
                        iden3_op + "<i>"             + section2_page1_3_9 + "</i>" + iden_ed, 
                        iden3_op + bold_blue_ital_op + section2_page1_3_10 + bold_blue_ital_ed + iden_ed]
    ##Page2
    section2_page2_1_1 = "The AMASS application de−duplicated the data by including only the first isolate per patient per specimen type per evaluation period as described in the method. " + \
                        "The number of patients with positive samples is as follows:"
    section2_page2_2_1 = "*The negative culture included data values specified as 'no growth' in the dictionary_for_microbiology_data file (details on data dictionary files are in the method section) to represent specimens with negative culture for any microorganism. "
    section2_page2_2_2 = "**Only the first isolate for each patient per specimen type, per pathogen, and per evaluation period was included in the analysis."
    section2_page2_3_1 = "The following figures and tables show the proportion of patients with blood culture positive for antimicrobial non−susceptible isolates."
    section2_page2_1 = [section2_page2_1_1]
    section2_page2_2 = [green_op + section2_page2_2_1 + green_ed, 
                        green_op + section2_page2_2_2 + green_ed]
    section2_page2_3 = [section2_page2_3_1]
    ##Page3-7
    section2_note_1 = "*Proportion of non−susceptible (NS) isolates represents the number of patients with blood culture positive for non−susceptible isolates (numerator) over the total number of patients with blood culture positive for the organism and the organism was tested for susceptibility against the antibiotic (denominator). " + \
                    "The AMASS application de−duplicated the data by including only the first isolate per patient per specimen type per evaluation period. Grey bars indicate that testing with the antibiotic occurred for less than 70% of the total number of patients with blood culture positive for the organism. "
    section2_note_2_1 = "CI=confidence interval; NA=Not available/reported/tested; "
    section2_note_2_2 = "CI=confidence interval; NA=Not available/reported/tested; 3GC=3rd−generation cephalosporin; "
    section2_note_2_drug1 = "Methicillin: methicillin, oxacillin, or cefoxitin"
    section2_note_2_drug2 = "FLUOROQUINOLONES: ciprofloxacin or levofloxacin; CARBAPENEMS: imipenem, meropenem, ertapenem or doripenem"
    section2_note_2_drug3 = "AMINOGLYCOSIDES: either gentamicin or amikacin; CARBAPENEMS: imipenem, meropenem, ertapenem or doripenem"
    section2_note = [section2_note_1 + section2_note_2_1 + section2_note_2_drug1]
    section2_notev2 = [section2_note_1 + section2_note_2_2 + section2_note_2_drug2]
    section2_notev3 = [section2_note_1 + section2_note_2_1 + section2_note_2_drug3]
    ######### SECTION2: PAGE1 #########
    report_title(c,'Section [2]: Isolate−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,section2_page1_1, 1.0*inch, 6.5*inch, 460, 200, font_size=11)
    report_title(c,'Organisms under this survey:',1.07*inch, 6.7*inch,'#3e4444',font_size=12)
    report_context(c,section2_page1_2, 1.0*inch, 3.8*inch, 460, 200, font_size=11)
    report_title(c,'Results',1.07*inch, 4.0*inch,'#3e4444',font_size=12)
    report_context(c,section2_page1_3, 1.0*inch, 0.2*inch, 460, 270, font_size=11, font_align=TA_LEFT)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()
    ######### SECTION2: PAGE2 #########
    report_context(c,section2_page2_1, 1.0*inch, 8.7*inch, 460, 100, font_size=11)
    report_context(c,section2_page2_2, 1.0*inch, 3.0*inch, 460, 120, font_size=11)
    report_context(c,section2_page2_3, 1.0*inch, 1.8*inch, 460, 50, font_size=11)
    table_draw = report2_table(summary_table)
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.0*inch, 5.5*inch)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.showPage()
    ######### SECTION2: PAGE3-7 #########
    d_head = {1:[lst_org_format[0],1.0*inch,9.5*inch,300,50, 
                 lst_numpat[0],    5.2*inch,9.5*inch,200,50],
              2:[lst_org_format[1],1.0*inch,5.5*inch,300,50, 
                 lst_numpat[1],    5.2*inch,5.5*inch,200,50],
              3:[lst_org_format[2],1.0*inch,9.5*inch,300,50, 
                 lst_numpat[2],    5.2*inch,9.5*inch,200,50],
              4:[lst_org_format[3],1.0*inch,5.5*inch,300,50, 
                 lst_numpat[3],    5.2*inch,5.5*inch,200,50],
              5:[lst_org_format[4],1.0*inch,9.5*inch,300,50, 
                 lst_numpat[4],    5.2*inch,9.5*inch,200,50],
              6:[lst_org_format[5],1.0*inch,9.5*inch,300,50, 
                 lst_numpat[5],    5.2*inch,9.5*inch,200,50],
              7:[lst_org_format[6],1.0*inch,9.5*inch,300,50, 
                 lst_numpat[6],    5.2*inch,9.5*inch,200,50],
              8:[lst_org_format[7],1.0*inch,5.5*inch,300,50, 
                 lst_numpat[7],    5.2*inch,5.5*inch,200,50]}
    d_pic = {1:[lst_org_short[0],0.8*inch,6.5*inch,3.2*inch,3.7*inch],
             2:[lst_org_short[1],0.8*inch,2.0*inch,3.3*inch,5.0*inch],
             3:[lst_org_short[2],0.8*inch,6.5*inch,3.2*inch,3.7*inch],
             4:[lst_org_short[3],0.5*inch,1.4*inch,3.3*inch,5.0*inch],
             5:[lst_org_short[4],0.4*inch,4.5*inch,3.3*inch,5.0*inch],
             6:[lst_org_short[5],0.4*inch,4.5*inch,3.3*inch,5.0*inch],
             7:[lst_org_short[6],0.4*inch,6.5*inch,3.2*inch,3.7*inch],
             8:[lst_org_short[7],0.4*inch,2.0*inch,3.3*inch,5.0*inch]}
    d_table = {1:[org1,500,300,4.4*inch,8.0*inch],
               2:[org2,500,300,4.4*inch,3.7*inch],
               3:[org3,500,300,4.4*inch,6.8*inch],
               4:[org4,500,300,4.0*inch,2.2*inch],
               5:[org5,500,300,4.0*inch,4.2*inch],
               6:[org6,500,300,4.0*inch,4.2*inch],
               7:[org7,500,300,4.0*inch,6.5*inch],
               8:[org8,500,300,4.0*inch,2.7*inch]}
    d_note = {1:["no_note"],
              2:[section2_note,1.0*inch,0.4*inch,460,120,lst_pagenumber[2],lastpage],
              3:["no_note"],
              4:[section2_notev2,1.0*inch,0.4*inch,460,120,lst_pagenumber[3],lastpage],
              5:[section2_notev2,1.0*inch,0.4*inch,460,120,lst_pagenumber[4],lastpage],
              6:[section2_notev2,1.0*inch,0.4*inch,460,120,lst_pagenumber[5],lastpage],
              7:["no_note"],
              8:[section2_notev3,1.0*inch,0.4*inch,460,120,lst_pagenumber[6],lastpage]}
    for keys in d_head.keys():
        report_title(c,'Section [2]: Isolate−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
        line_1_1 = ["<b>" + "Blood: " + d_head[keys][0] + "</b>"]
        line_1_2 = [bold_blue_ital_op + "( No. of patients = " + str(d_head[keys][5]) + " )" + bold_blue_ital_ed]
        report_context(c,line_1_1, d_head[keys][1], d_head[keys][2], d_head[keys][3], d_head[keys][4], font_size=12, font_align=TA_LEFT)
        report_context(c,line_1_2, d_head[keys][6], d_head[keys][7], d_head[keys][8], d_head[keys][9], font_size=12, font_align=TA_LEFT)
        c.drawImage(path_result + 'Report2_AMR_' + d_pic[keys][0] +".png", d_pic[keys][1], d_pic[keys][2], preserveAspectRatio=True, width=d_pic[keys][3], height=d_pic[keys][4],showBoundary=False) 
        table_draw = report2_table_nons(d_table[keys][0])
        table_draw.wrapOn(c, d_table[keys][1], d_table[keys][2])
        table_draw.drawOn(c, d_table[keys][3], d_table[keys][4])
        if d_note[keys][0] != "no_note":
            report_context(c,d_note[keys][0], d_note[keys][1], d_note[keys][2], d_note[keys][3], d_note[keys][4], font_size=9,line_space=12)
            report_todaypage(c,55,30,"Created on: "+today)
            report_todaypage(c,270,30,"Page " + d_note[keys][5] + " of " + d_note[keys][6])
            c.showPage()
        else:
            pass

def section3(result_table, summary_table, 
             org1_com, org1_hos, org2_com, org2_hos, org3_com, org3_hos, org4_com, org4_hos, 
             org5_com, org5_hos, org6_com, org6_hos, org7_com, org7_hos, org8_com, org8_hos, 
             lst_org, lst_numpat, lst_org_short, lst_pagenumber=pagenumber_ava_3,  lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variable
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start  = str(sec3_res.loc[sec3_res["Parameters"]=="Minimum_date","Values"].tolist()[0])
    spc_date_end    = str(sec3_res.loc[sec3_res["Parameters"]=="Maximum_date","Values"].tolist()[0])
    pat_num_pos_org     = str(sec3_res.loc[sec3_res["Parameters"]=="Number_of_patients_with_blood_culture_positive_for_organism_under_this_survey","Values"].tolist()[0])
    pat_num_pos_org_com = str(sec3_res.loc[sec3_res["Parameters"]=="Number_of_patients_with_community_origin_BSI","Values"].tolist()[0])
    pat_num_pos_org_hos = str(sec3_res.loc[sec3_res["Parameters"]=="Number_of_patients_with_hospital_origin_BSI","Values"].tolist()[0])
    pat_num_pos_org_unk = str(sec3_res.loc[sec3_res["Parameters"]=="Number_of_patients_with_unknown_origin_BSI","Values"].tolist()[0])
    ##Page1
    section3_page1_1_1 = "An isolate−based surveillance report with stratification by origin of infection is generated only if admission date data are available in the raw data file(s) with the appropriate specification in the data dictionaries."
    section3_page1_1_2 = "Stratification by origin of infection is used as a proxy to define where the bloodstream infection (BSI) was contracted (hospital versus community)."
    section3_page1_1_3 = "The definitions of infection origin proposed by the WHO GLASS are used. In brief, community−origin BSI is defined as patients in the hospital for less than or equal to two calendar days when the first specimen culture postive for the pathogen was taken. " + \
                        "Hospital−origin BSI is defined as patients admitted for more than two calendar days when the first specimen culture positive for the pathogen was taken."
    section3_page1_1 = [section3_page1_1_1, 
                        add_blankline + section3_page1_1_2, 
                        add_blankline + section3_page1_1_3]
    section3_page1_2_1 = "The data included in the analysis to generate the report had:"
    section3_page1_2_2 = "<i>" + "Sample collection dates ranged from " + "</i>" + \
                        bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + \
                        "<i>" + " to " + "</i>" +\
                        bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section3_page1_2_3 = "*Number of patients with blood culture positive for pathogen under the survey:"
    section3_page1_2_4 = pat_num_pos_org + " patients"
    section3_page1_2_5 = "**Number of patients with community−origin BSI:"
    section3_page1_2_6 = pat_num_pos_org_com + " patients"
    section3_page1_2_7 = "**Number of patients with hospital−origin BSI:"
    section3_page1_2_8 = pat_num_pos_org_hos + " patients"
    section3_page1_2_9 = "***Number of patients with unknown infection of origin status:"
    section3_page1_2_10 = pat_num_pos_org_unk + " patients"
    section3_page1_2 = [section3_page1_2_1, 
                        iden1_op + section3_page1_2_2 + iden_ed, 
                        iden1_op +"<i>" + section3_page1_2_3 + "</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section3_page1_2_4 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>" + section3_page1_2_5 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section3_page1_2_6 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>" + section3_page1_2_7 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section3_page1_2_8 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>" + section3_page1_2_9 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section3_page1_2_10 + bold_blue_ital_ed + iden_ed]
    ##Page2
    section3_page2_org = ["<b>" + "Organism" + "</b>", 
                        "<b>" + "" + "</b>", 
                        "<b>" + "" + "</b>", 
                        "<b>" + "" + "</b>", 
                        "<b>" + lst_org_format[0] + "</b>", 
                        "<b>" + lst_org_format[1] + "</b>", 
                        "<b>" + lst_org_format[2] + "</b>", 
                        "<b>" + lst_org_format[3] + "</b>", 
                        "<b>" + lst_org_format[4] + "</b>", 
                        "<b>" + lst_org_format[5] + "</b>", 
                        "<b>" + lst_org_format[6] + "</b>", 
                        "<b>" + lst_org_format[7] + "</b>", 
                        "<b>" + "Total:" + "</b>"]
    section3_page2_1_1 = "NA= Not applicable (hospital admission date or infection origin data are not available)"
    section3_page2_1_2 = "*Only the first isolate for each patient per specimen type per pathogen under the reporting period is included in the analysis. Please refer to Section [2] for details on how this number was calculated from the raw microbiology_data file."
    section3_page2_1_3 = "**The definitions of infection origin proposed by the WHO GLASS is used. In brief, community−origin BSI was defined as patients in the hospital for less than or equal to two calendar days when the first blood culture positive for the pathogen was taken."
    section3_page2_1_4 = "Hospital−origin BSI was defined as patients admitted for more than two calendar days when the first specimen culture positive for the pathogen was taken."
    section3_page2_1_5 = "Please refer to the 'Methods' section for more details on the definitions used."
    section3_page2_1_6 = "***Unknown origin could be because admission date data are not available or the patient was not hospitalised."
    section3_page2_1 = [green_op + section3_page2_1_1 + green_ed, 
                        green_op + add_blankline + section3_page2_1_2 + green_ed, 
                        green_op + section3_page2_1_3 + green_ed, 
                        green_op + section3_page2_1_4 + green_ed, 
                        green_op + section3_page2_1_5 + green_ed, 
                        green_op + section3_page2_1_6 + green_ed]
    section3_page2_2_1 = "The following figures and tables below show the proportion of patients with blood culture positive for antimicrobial non−susceptible isolates stratified by infection of origin."
    section3_page2_2 = [section3_page2_2_1]
    section2_note_1 = "*Proportion of non−susceptible (NS) isolates represents the number of patients with blood culture positive for non−susceptible isolates (numerator) over the total number of patients with blood culture positive for the organism and the organism was tested for susceptibility against the antibiotic (denominator). " + \
                    "The AMASS application de−duplicated the data by including only the first isolate per patient per specimen type per evaluation period. Grey bars indicate that testing with the antibiotic occurred for less than 70% of the total number of patients with blood culture positive for the organism. "
    section2_note_2_1 = "CI=confidence interval; NA=Not available/reported/tested; "
    section2_note_2_2 = "CI=confidence interval; NA=Not available/reported/tested; 3GC=3rd−generation cephalosporin; "
    section2_note_2_drug1 = "Methicillin: methicillin, oxacillin, or cefoxitin "
    section2_note_2_drug2 = "FLUOROQUINOLONES: ciprofloxacin or levofloxacin; CARBAPENEMS: imipenem, meropenem, ertapenem or doripenem"
    section2_note_2_drug3 = "AMINOGLYCOSIDES: either gentamicin or amikacin; CARBAPENEMS: imipenem, meropenem, ertapenem or doripenem"
    section2_note = [section2_note_1 + section2_note_2_1 + section2_note_2_drug1]
    section2_notev2 = [section2_note_1 + section2_note_2_2 + section2_note_2_drug2]
    section2_notev3 = [section2_note_1 + section2_note_2_1 + section2_note_2_drug3]
    ######### SECTION3: PAGE1 #########
    report_title(c,'Section [3]: Isolate−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,section3_page1_1, 1.0*inch, 5.9*inch, 460, 250, font_size=11)
    report_title(c,'Results',1.07*inch, 5.7*inch,'#3e4444',font_size=12)
    report_context(c,section3_page1_2, 1.0*inch, 2.8*inch, 460, 200, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()
    ######### SECTION3: PAGE2 #########
    report_title(c,'Note',1.07*inch, 6.0*inch,'darkgreen',font_size=12)
    report_context(c,section3_page2_1, 1.0*inch, 2.5*inch, 460, 250, font_size=11)
    report_context(c,section3_page2_2, 1.0*inch, 1.5*inch, 460, 50, font_size=11)
    report_context(c,section3_page2_org, 1.0*inch, 6.0*inch, 460, 250, font_size=11, line_space=18.5)
    table_draw = report3_table(sec3_pat_val)
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 3.2*inch, 6.9*inch)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.showPage()
    ######### SECTION3: PAGE3-12 #########
    d_head = {"com_1":[lst_org_format[0], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, lst_numpat[0], 5.5*inch, 9.0*inch],
              "hos_1":[lst_org_format[0], 1.0*inch, 5.5*inch, 4.0*inch, 5.5*inch, lst_numpat[1], 5.5*inch, 5.5*inch],
              "com_2":[lst_org_format[1], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, lst_numpat[2], 5.5*inch, 9.0*inch],
              "hos_2":[lst_org_format[1], 1.0*inch, 5.5*inch, 4.0*inch, 5.5*inch, lst_numpat[3], 5.5*inch, 5.5*inch],
              "com_3":[lst_org_format[2], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, lst_numpat[4], 5.5*inch, 9.0*inch],
              "hos_3":[lst_org_format[2], 1.0*inch, 5.5*inch, 4.0*inch, 5.5*inch, lst_numpat[5], 5.5*inch, 5.5*inch],
              "com_4":[lst_org_format[3], 1.0*inch, 9.3*inch, 4.0*inch, 9.3*inch, lst_numpat[6], 5.5*inch, 9.3*inch],
              "hos_4":[lst_org_format[3], 1.0*inch, 5.3*inch, 4.0*inch, 5.3*inch, lst_numpat[7], 5.5*inch, 5.3*inch],
              "com_5":[lst_org_format[4], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, lst_numpat[8], 5.5*inch, 9.0*inch],
              "hos_5":[lst_org_format[4], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, lst_numpat[9], 5.5*inch, 9.0*inch],
              "com_6":[lst_org_format[5], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, lst_numpat[10], 5.5*inch, 9.0*inch],
              "hos_6":[lst_org_format[5], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, lst_numpat[11], 5.5*inch, 9.0*inch],
              "com_7":[lst_org_format[6], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, lst_numpat[12], 5.5*inch, 9.0*inch],
              "hos_7":[lst_org_format[6], 1.0*inch, 5.2*inch, 4.0*inch, 5.2*inch, lst_numpat[13], 5.5*inch, 5.2*inch],
              "com_8":[lst_org_format[7], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, lst_numpat[14], 5.5*inch, 9.0*inch],
              "hos_8":[lst_org_format[7], 1.0*inch, 5.2*inch, 4.0*inch, 5.2*inch, lst_numpat[15], 5.5*inch, 5.2*inch]}
    d_pic =  {"com_1":[lst_org_short[0], 0.8*inch, 6.2*inch, 3.2*inch, 3.7*inch, "Community"],
              "hos_1":[lst_org_short[0], 0.8*inch, 2.0*inch, 3.3*inch, 5.0*inch, "Hospital"],
              "com_2":[lst_org_short[1], 0.8*inch, 6.2*inch, 3.2*inch, 3.7*inch, "Community"],
              "hos_2":[lst_org_short[1], 0.8*inch, 2.0*inch, 3.3*inch, 5.0*inch, "Hospital"],
              "com_3":[lst_org_short[2], 0.8*inch, 6.0*inch, 3.2*inch, 3.7*inch, "Community"],
              "hos_3":[lst_org_short[2], 0.8*inch, 1.8*inch, 3.3*inch, 5.0*inch, "Hospital"],
              "com_4":[lst_org_short[3], 0.4*inch, 6.3*inch, 3.2*inch, 3.7*inch, "Community"],
              "hos_4":[lst_org_short[3], 0.4*inch, 1.8*inch, 3.3*inch, 4.5*inch, "Hospital"],
              "com_5":[lst_org_short[4], 0.4*inch, 3.4*inch, 3.5*inch, 5.5*inch, "Community"],
              "hos_5":[lst_org_short[4], 0.4*inch, 3.4*inch, 3.5*inch, 5.5*inch, "Hospital"],
              "com_6":[lst_org_short[5], 0.4*inch, 3.4*inch, 3.5*inch, 5.5*inch, "Community"],
              "hos_6":[lst_org_short[5], 0.4*inch, 3.4*inch, 3.5*inch, 5.5*inch, "Hospital"],
              "com_7":[lst_org_short[6], 0.4*inch, 6.0*inch, 3.2*inch, 3.7*inch, "Community"],
              "hos_7":[lst_org_short[6], 0.4*inch, 1.5*inch, 3.3*inch, 5.0*inch, "Hospital"],
              "com_8":[lst_org_short[7], 0.4*inch, 6.0*inch, 3.2*inch, 3.7*inch, "Community"],
              "hos_8":[lst_org_short[7], 0.4*inch, 1.5*inch, 3.3*inch, 5.0*inch, "Hospital"]}
    d_table = {"com_1":[org1_com,500,300,4.4*inch,7.7*inch],
               "hos_1":[org1_hos,500,300,4.4*inch,4.0*inch],
               "com_2":[org2_com,500,300,4.4*inch,7.2*inch],
               "hos_2":[org2_hos,500,300,4.4*inch,3.8*inch],
               "com_3":[org3_com,500,300,4.4*inch,6.4*inch],
               "hos_3":[org3_hos,500,300,4.4*inch,2.8*inch],
               "com_4":[org4_com,500,300,4.0*inch,6.1*inch],
               "hos_4":[org4_hos,500,300,4.0*inch,2.2*inch],
               "com_5":[org5_com,500,300,4.0*inch,3.7*inch],
               "hos_5":[org5_hos,500,300,4.0*inch,3.7*inch],
               "com_6":[org6_com,500,300,4.0*inch,3.7*inch],
               "hos_6":[org6_hos,500,300,4.0*inch,3.7*inch],
               "com_7":[org7_com,500,300,4.0*inch,6.0*inch],
               "hos_7":[org7_hos,500,300,4.0*inch,2.20*inch],
               "com_8":[org8_com,500,300,4.0*inch,6.2*inch],
               "hos_8":[org8_hos,500,300,4.0*inch,2.4*inch]}
    d_note = {"com_1":["no_note"],
               "hos_1":[section2_note,1.0*inch,0.4*inch,460,120,lst_pagenumber[2],lastpage],
               "com_2":["no_note"],
               "hos_2":[section2_note,1.0*inch,0.4*inch,460,120,lst_pagenumber[3],lastpage],
               "com_3":["no_note"],
               "hos_3":[section2_notev2,1.0*inch,0.4*inch,460,120,lst_pagenumber[4],lastpage],
               "com_4":["no_note"],
               "hos_4":[section2_notev2,1.0*inch,0.4*inch,460,120,lst_pagenumber[5],lastpage],
               "com_5":[section2_notev2,1.0*inch,0.4*inch,460,120,lst_pagenumber[6],lastpage],
               "hos_5":[section2_notev2,1.0*inch,0.4*inch,460,120,lst_pagenumber[7],lastpage],
               "com_6":[section2_notev2,1.0*inch,0.4*inch,460,120,lst_pagenumber[8],lastpage],
               "hos_6":[section2_notev2,1.0*inch,0.4*inch,460,120,lst_pagenumber[9],lastpage],
               "com_7":["no_note"],
               "hos_7":[section2_notev3,1.0*inch,0.4*inch,460,120,lst_pagenumber[10],lastpage],
               "com_8":["no_note"],
               "hos_8":[section2_notev3,1.0*inch,0.4*inch,460,120,lst_pagenumber[11],lastpage]}
    for keys in d_head.keys():
        report_title(c,'Section [3]: Isolate−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
        report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
        line_1_1 = ["<b>" + "Blood: " + d_head[keys][0] + "</b>"]
        line_1_2 = ["<b>" + d_pic[keys][5] + "-origin" + "</b>"]
        line_1_3 = [bold_blue_ital_op + "( No. of patients = " + str(d_head[keys][5]) + " )" + bold_blue_ital_ed]
        report_context(c,line_1_1, d_head[keys][1], d_head[keys][2], 300, 50, font_size=12, font_align=TA_LEFT)
        report_context(c,line_1_2, d_head[keys][3], d_head[keys][4], 200, 50, font_size=12, font_align=TA_LEFT)
        report_context(c,line_1_3, d_head[keys][6], d_head[keys][7], 200, 50, font_size=12, font_align=TA_LEFT)
        c.drawImage(path_result + 'Report3_AMR_' + d_pic[keys][0] + "_" + d_pic[keys][5]+".png", d_pic[keys][1], d_pic[keys][2], preserveAspectRatio=True, width=d_pic[keys][3], height=d_pic[keys][4],showBoundary=False) 
        table_draw = report2_table_nons(d_table[keys][0])
        table_draw.wrapOn(c, d_table[keys][1], d_table[keys][2])
        table_draw.drawOn(c, d_table[keys][3], d_table[keys][4])
        if d_note[keys][0] != "no_note":
            report_context(c,d_note[keys][0], d_note[keys][1], d_note[keys][2], d_note[keys][3], d_note[keys][4], font_size=9,line_space=12)
            report_todaypage(c,55,30,"Created on: "+today)
            report_todaypage(c,270,30,"Page " + d_note[keys][5] + " of " + d_note[keys][6])
            c.showPage()
        else:
            pass

def section4(result_table, result_blo_table, result_pat_table,
             lst_pagenumber=pagenumber_ava_4, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variable
    iden1_op = "<para leftindent=\"35\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start = result_table.loc[result_table["Parameters"]=="Minimum_date","Values"].tolist()[0]
    spc_date_end = result_table.loc[result_table["Parameters"]=="Maximum_date","Values"].tolist()[0]
    blo_num = result_table.loc[result_table["Parameters"]=="Number_of_blood_specimens_collected","Values"].tolist()[0]
    pat_num_pos_blo = result_table.loc[result_table["Parameters"]=="Number_of_patients_sampled_for_blood_culture","Values"].tolist()[0]
    ##Page1
    section4_page1_1_1 = "A sample−based surveillance report is generated if data of culture negative is available."
    section4_page1_1_2 = "The sample−based approach involves the collection of data on all blood samples taken for microbiological testing and includes information on the number of positive blood samples for a specific specimen type (both pathogens under the survey and other bacteria) as well as number of negative (no microbial growth) samples. " + \
                        "After removal of duplicate results and assuming that routine blood culture testing is applied systematically, we can use the number of tested patients as a proxy for a number of patients with new cases of bloodstream infection (BSI)."
    section4_page1_1 = [section4_page1_1_1, add_blankline + section4_page1_1_2]
    section4_page1_2_1 = "The microbiology_data file had:"
    section4_page1_2_2 = "<i>" + "Specimen collection dates ranged from " +"</i>" +\
                        bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + \
                        "<i>" + " to " + "</i>" + \
                        bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section4_page1_2_3 = "Number of records on blood specimen collected within the above date range:"
    section4_page1_2_4 = blo_num +" blood specimen records"
    section4_page1_2_5 = "*Number of patients sampled for blood culture within the above date range:"
    section4_page1_2_6 = pat_num_pos_blo + " patients sampled for blood culture"
    section4_page1_2 = [section4_page1_2_1, 
                        iden1_op + section4_page1_2_2 + iden_ed, 
                        iden1_op + "<i>" + section4_page1_2_3 +"</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op +  section4_page1_2_4 + bold_blue_ital_ed + iden_ed, 
                        iden1_op + "<i>" + section4_page1_2_5 +"</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section4_page1_2_6 + bold_blue_ital_ed + iden_ed]
    section4_page1_3_1 = "*Number of patients sampled for blood culture is used as denominator to estimate the frequency of infections per 100,000 tested patients"
    section4_page1_3 = [green_op + section4_page1_3_1 + green_ed]
    section4_page1_4_1 = "The following figures show the frequncy of infections for patients with blood culture tested."
    section4_page1_4 = [section4_page1_4_1]
    ##Page2-3
    section4_page2_1 = "*Frequency of infection per 100,000 tested patients represents the number of patients with blood culture positive for a pathogen (numerator) over the total number of tested patients (denominator). " + \
                        "The AMASS application de−duplicates the data by included only the first isolate of each patient per specimen type per reporting period."
    section4_page2_2 = "CI=confidence interval; NS=non−susceptible; NA=Not available/reported/tested; 3GC=3rd−generation cephalosporin"
    section4_page2 = [section4_page2_1, section4_page2_2]
    ######### SECTION4: PAGE1 #########
    report_title(c,'Section [4]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,section4_page1_1, 1.0*inch, 6.6*inch, 460, 200, font_size=11)
    report_title(c,'Results',1.07*inch, 6.5*inch,'#3e4444',font_size=12)
    report_context(c,section4_page1_2, 1.0*inch, 4.3*inch, 460, 150, font_size=11)
    report_title(c,'Note',1.07*inch, 3.5*inch,'darkgreen',font_size=12)
    report_context(c,section4_page1_3, 1.0*inch, 2.7*inch, 460, 50, font_size=11)
    report_context(c,section4_page1_4, 1.0*inch, 1.5*inch, 460, 50, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()
    ######### SECTION4: PAGE2 #########
    report_title(c,'Section [4]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Pathogens" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_pos_blo) + " )" + bold_blue_ital_ed]
    report_context(c,line_1_1, 1.0*inch, 9.3*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_3, 5.5*inch, 9.3*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    c.drawImage(path_result+"Report4_frequency_blood.png", 0.7*inch, 3.0*inch, preserveAspectRatio=False, width=3.5*inch, height=6.5*inch,showBoundary=False) 
    table_draw = report2_table_nons(result_blo_table)
    table_draw.wrapOn(c, 230, 300)
    table_draw.drawOn(c, 4.3*inch, 5.5*inch)
    report_context(c,section4_page2, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.showPage()
    ######### SECTION4: PAGE3 #########
    report_title(c,'Section [4]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Non-susceptible pathogens" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_pos_blo) + " )" + bold_blue_ital_ed]
    report_context(c,line_1_1, 1.0*inch, 9.3*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_3, 5.5*inch, 9.3*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    c.drawImage(path_result+"Report4_frequency_pathogen.png", 0.7*inch, 3.0*inch, preserveAspectRatio=False, width=3.5*inch, height=6.5*inch,showBoundary=False) 
    table_draw = report2_table_nons(result_pat_table)
    table_draw.wrapOn(c, 240, 300)
    table_draw.drawOn(c, 4.3*inch, 4.5*inch)
    report_context(c,section4_page2, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[2] + " of " + lastpage)
    c.showPage()

def section5(result_table, result_com_table, result_hos_table, result_com_amr_table, result_hos_amr_table,
             lst_pagenumber=pagenumber_ava_5, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start     = result_table.loc[result_table["Parameters"]=="Minimum_date","Values"].tolist()[0]
    spc_date_end     = result_table.loc[result_table["Parameters"]=="Maximum_date","Values"].tolist()[0]
    blo_num     = result_table.loc[result_table["Parameters"]=="Number_of_blood_specimens_collected","Values"].tolist()[0]
    pat_num_pos_blo     = result_table.loc[result_table["Parameters"]=="Number_of_patients_sampled_for_blood_culture","Values"].tolist()[0]
    pat_num_w2day   = result_table.loc[result_table["Parameters"]=="Number_of_patients_with_blood_culture_within_first_2_days_of_admission","Values"].tolist()[0]
    pat_num_wo2day  = result_table.loc[result_table["Parameters"]=="Number_of_patients_with_blood_culture_within_after_2_days_of_admission","Values"].tolist()[0]
    pat_num_unk     = result_table.loc[result_table["Parameters"]=="Number_of_patients_with_unknown_origin","Values"].tolist()[0]
    pat_num_oth     = result_table.loc[result_table["Parameters"]=="Number_of_patients_had_more_than_one_admission","Values"].tolist()[0]
    ##Page1
    section5_page1_1_1 = "A sample−based surveillance report with stratification by origin of infection is generated only if data of culture negative is available and admission date or a variable containing the classification is available in the raw data file with the appropriate specification in the data dictionaries."
    section5_page1_1 = [section5_page1_1_1]
    section5_page1_2_1 = "The data included in the analysis had:"
    section5_page1_2_2 = "<i>" + "Specimen collection dates ranged from " + "</i>" + \
                        bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + \
                        " to " + \
                        bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section5_page1_2_3 = "Number of records on blood specimen collected within the above date range:"
    section5_page1_2_4 = blo_num + " blood specimen records"
    section5_page1_2_5 = "Number of patients sampled for blood culture within the above date range:"
    section5_page1_2_6 = pat_num_pos_blo + " patients sampled for blood culture"
    section5_page1_2_7 = bold_blue_ital_op + pat_num_w2day + bold_blue_ital_ed + \
                        "<i>" + " patients had at least one admission having the first blood culture drawn within first 2 calendar days of hospital admission." + "</i>"
    section5_page1_2_8 = "This parameter is used as a denominators for frequency of community−origin bacteraemia (per 100,000 patients tested for blood culture on admission)."
    section5_page1_2_9 = bold_blue_ital_op + pat_num_wo2day + bold_blue_ital_ed + \
                        "<i>" + " patients had at least one admission having the first blood culture drawn after 2 calendar days of hospital admission." + "</i>"
    section5_page1_2_10 = "This parameter is used as a denominators for frequency of hospital−origin bacteraemia (per 100,000 patients tested for blood culture for HAI)."
    section5_page1_2_11 = bold_blue_ital_op + pat_num_unk + bold_blue_ital_ed + \
                        "<i>" + " patients had a blood drawn for culture and with unknown origin of infection." + "</i>"
    section5_page1_2_12 = "Validation of this statistics is highly recommended."
    section5_page1_2 = [section5_page1_2_1, 
                        iden1_op + section5_page1_2_2 + iden_ed, 
                        iden1_op + "<i>" + section5_page1_2_3+ "</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section5_page1_2_4 + bold_blue_ital_ed + iden_ed, 
                        iden1_op + "<i>" + section5_page1_2_5+ "</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section5_page1_2_6 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + add_blankline + section5_page1_2_7 + iden_ed, 
                        iden2_op + "<i>" + section5_page1_2_8 + "</i>" + iden_ed, 
                        iden2_op + section5_page1_2_9 + iden_ed, 
                        iden2_op + "<i>" + section5_page1_2_10 + "</i>" + iden_ed, 
                        iden2_op + section5_page1_2_11 + iden_ed, 
                        iden2_op + "<i>" + section5_page1_2_12 + "</i>" + iden_ed]
    section5_page1_3_1 = bold_blue_ital_op + pat_num_oth + bold_blue_ital_ed + \
                        "<i>" + green_op + " patients had more than one admissions, of which at least one admission had the first blood culture drawn within the first 2 calendar days of hospital admission AND at least one admission had the first blood culture drawn after 2 calendar days of hospital admission." + \
                        green_ed + "</i>"
    section5_page1_3 = [iden2_op + section5_page1_3_1 + iden_ed]
    section5_page1_4_1 = "The following figures show the frequency of infections for patients with blood culture tested and stratified by infection origin, under this surveillance."
    section5_page1_4 = [section5_page1_4_1]
    ##Page2-5
    section5_page2_1 = "*Frequency of infection per 100,000 tested patients on admission represents the number of patients with blood culture positive for a pathogen (numerator) over the total number of tested population on admission (denominator). " + \
                        "The AMASS application de−duplicates the data by included only the first isolate of each patient per specimen type per reporting period."
    section5_page2_1v2= "*Frequency of infection per 100,000 tested population at risk of HAI represents the number of patients with blood culture positive for a pathogen (numerator) over the total number of tested population at risk of HAI (denominator). The AMASS application de−duplicates the data by included only the first isolate of each patient per specimen type per reporting period."
    section5_page2_1v3= "*Frequency of infection per 100,000 tested patients represents the number of patients with blood culture positive for a pathogen (numerator) over the total number of tested patients (denominator). The AMASS application de−duplicates the data by included only the first isolate of each patient per specimen type per reporting period."
    section5_page2_2 = "CI=confidence interval; NS=non−susceptible; NA=Not available/reported/tested; 3GC=3rd−generation cephalosporin"
    section5_page2 = [section5_page2_1, section5_page2_2]
    section5_page2v2 = [section5_page2_1v2, section5_page2_2]
    section5_page2v3 = [section5_page2_1v3, section5_page2_2]
    ######### SECTION5: PAGE1 #########
    report_title(c,'Section [5]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,section5_page1_1, 1.0*inch, 8.0*inch, 460, 100, font_size=11)
    report_title(c,'Results',1.07*inch, 7.9*inch,'#3e4444',font_size=12)
    report_context(c,section5_page1_2, 1.0*inch, 3.2*inch, 480, 330, font_size=11)
    report_title(c,'Note:',2.0*inch, 3.0*inch,'darkgreen',font_size=12)
    report_context(c,section5_page1_3, 1.0*inch, 1.5*inch, 460, 100, font_size=11)
    report_context(c,section5_page1_4, 1.0*inch, 1.0*inch, 460, 50, font_size=11)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()
    ######### SECTION5: PAGE2 #########
    report_title(c,'Section [5]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Pathogens" + "</b>"]
    line_1_2 = ["<b>" + "Community-origin" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_w2day) + " )" + bold_blue_ital_ed]
    report_context(c,line_1_1, 1.0*inch, 9.0*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_2, 4.0*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_3, 5.5*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    c.drawImage(path_result+"Report5_incidence_community.png", 0.7*inch, 2.7*inch, preserveAspectRatio=False, width=3.5*inch, height=6.5*inch,showBoundary=False) 
    table_draw = report2_table_nons(result_com_table)
    table_draw.wrapOn(c, 230, 300)
    table_draw.drawOn(c, 4.3*inch, 5.2*inch)
    report_context(c,section5_page2, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.showPage()
    ######### SECTION5: PAGE3 #########
    report_title(c,'Section [5]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Pathogens" + "</b>"]
    line_1_2 = ["<b>" + "Hospital-origin" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_wo2day) + " )" + bold_blue_ital_ed]
    report_context(c,line_1_1, 1.0*inch, 9.0*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_2, 4.0*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_3, 5.5*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    c.drawImage(path_result+"Report5_incidence_hospital.png", 0.7*inch, 2.7*inch, preserveAspectRatio=False, width=3.5*inch, height=6.5*inch,showBoundary=False) 
    table_draw = report2_table_nons(result_hos_table)
    table_draw.wrapOn(c, 230, 300)
    table_draw.drawOn(c, 4.3*inch, 5.2*inch)
    report_context(c,section5_page2v2, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[2] + " of " + lastpage)
    c.showPage()
    ######### SECTION5: PAGE4 #########
    report_title(c,'Section [5]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Non-susceptible pathogens" + "</b>"]
    line_1_2 = ["<b>" + "Community-origin" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_w2day) + " )" + bold_blue_ital_ed]
    report_context(c,line_1_1, 1.0*inch, 9.0*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_2, 4.0*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_3, 5.5*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    c.drawImage(path_result+"Report5_incidence_community_antibiotic.png", 0.7*inch, 2.7*inch, preserveAspectRatio=False, width=3.5*inch, height=6.5*inch,showBoundary=False) 
    table_draw = report2_table_nons(result_com_amr_table)
    table_draw.wrapOn(c, 240, 300)
    table_draw.drawOn(c, 4.3*inch, 4.3*inch)
    report_context(c,section5_page2, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[3] + " of " + lastpage)
    c.showPage()
    ######### SECTION5: PAGE5 #########
    report_title(c,'Section [5]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Non-susceptible pathogens" + "</b>"]
    line_1_2 = ["<b>" + "Hospital-origin" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_wo2day) + " )" + bold_blue_ital_ed]
    report_context(c,line_1_1, 1.0*inch, 9.0*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_2, 4.0*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    report_context(c,line_1_3, 5.5*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    c.drawImage(path_result+"Report5_incidence_hospital_antibiotic.png", 0.7*inch, 2.7*inch, preserveAspectRatio=False, width=3.5*inch, height=6.5*inch,showBoundary=False) 
    table_draw = report2_table_nons(result_hos_amr_table)
    table_draw.wrapOn(c, 240, 300)
    table_draw.drawOn(c, 4.3*inch, 4.3*inch)
    report_context(c,section5_page2v3, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[4] + " of " + lastpage)
    c.showPage()

def section6(result_table, result_mor_table, 
             org1_com, org1_hos, org2_com, org2_hos, org3_com, org3_hos, org4_com, org4_hos, 
             org5_com, org5_hos, org6_com, org6_hos, org7_com, org7_hos, org8_com, org8_hos,
             lst_org, lst_org_short, lst_org_full, df_numpat_com, df_numpat_hos,
             lst_pagenumber=pagenumber_ava_6, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden3_op = "<para leftindent=\"105\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    tab1st = "&nbsp;"
    tab4th = "&nbsp;&nbsp;&nbsp;&nbsp;"
    ##variables
    spc_date_start = result_table.loc[(result_table["Type_of_data_file"]=="microbiology_data")&(result_table["Parameters"]=="Minimum_date"),"Values"].tolist()[0]
    spc_date_end       = result_table.loc[(result_table["Type_of_data_file"]=="microbiology_data")&(result_table["Parameters"]=="Maximum_date"),"Values"].tolist()[0]
    pat_num_pos_org    = result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_positive_for_organism_under_this_survey","Values"].tolist()[0]
    pat_num_pos_org_com= result_table.loc[result_table["Parameters"]=="Number_of_patients_sampled_for_blood_culture","Values"].tolist()[0]
    pat_num_pos_org_hos= result_table.loc[result_table["Parameters"]=="Number_of_patients_with_community_origin_BSI","Values"].tolist()[0]
    hos_date_start     = result_table.loc[(result_table["Type_of_data_file"]=="hospital_admission_data")&(result_table["Parameters"]=="Minimum_date"),"Values"].tolist()[0]
    hos_date_end       = result_table.loc[(result_table["Type_of_data_file"]=="hospital_admission_data")&(result_table["Parameters"]=="Maximum_date"),"Values"].tolist()[0]
    hos_num            = result_table.loc[result_table["Parameters"]=="Number_of_records","Values"].tolist()[0]
    pat_num_hos        = result_table.loc[result_table["Parameters"]=="Number_of_patients_included","Values"].tolist()[0]
    pat_num_dead       = result_table.loc[result_table["Parameters"]=="Number_of_deaths","Values"].tolist()[0]
    per_mortal         = result_table.loc[result_table["Parameters"]=="Mortality","Values"].tolist()[0]
    ##Page1
    section6_page1_1_1 = "A surveillance report on mortality involving AMR infections and antimicrobial−susceptible infections with stratification by origin of infection is generated only if data on patient outcomes (i.e. discharge status) are available. " + \
                    "Antimicrobial−resistant infection is a threat to modern health care, and the impact of the infection on patient outcomes is largely unknown. " + \
                    "Performing analyses and generating reports on mortality often takes time and resources."
    section6_page1_1_2 = "The term 'mortality involving AMR and antimicrobial−susceptible infections was used because the mortality reported was all−cause mortality. " + \
                        "This measure of mortality included deaths caused by or related to other underlying and intermediate causes."
    section6_page1_1_3 = "Here, AMASS summarized the overall mortality of patients with antimicrobial−resistant and antimicrobial−susceptible bacteria bloodstream infections (BSI)."
    section6_page1_1 = [section6_page1_1_1, 
                        add_blankline + section6_page1_1_2, 
                        add_blankline + section6_page1_1_3]
    section6_page1_2_1 = "The data included in the analysis had:"
    section6_page1_2_2 = "Sample collection dates ranged from "+ \
                        bold_blue_ital_op + str(spc_date_start) + bold_blue_ital_ed + \
                        "  to  " + \
                        bold_blue_ital_op + str(spc_date_end) + bold_blue_ital_ed
    section6_page1_2_3 = "Number of patients with blood culture positive for the origanism under the survey:"
    section6_page1_2_4 = str(pat_num_pos_org) + " patients"
    section6_page1_2_5 = "Number of patients with community−origin BSI:"
    section6_page1_2_6 = str(pat_num_pos_org_com) + " patients"
    section6_page1_2_7 = "Number of patients with hospital−origin BSI:"
    section6_page1_2_8 = str(pat_num_pos_org_hos) + " patients"
    section6_page1_2_9 = "The hospital admission data file had:"
    section6_page1_2_10 = "Hospital admission dates ranging from "+ \
                        bold_blue_ital_op + str(hos_date_start) + bold_blue_ital_ed + \
                        "  to  " + \
                        bold_blue_ital_op + str(hos_date_end) + bold_blue_ital_ed
    section6_page1_2_11 = "Number of records in the raw hospital admission data:"
    section6_page1_2_12 = str(hos_num) + " records"
    section6_page1_2_13 = "Number of patients included in the analysis (de−duplicated):"
    section6_page1_2_14 = str(pat_num_hos) + " patients"
    section6_page1_2_15 = "Number of patients having death as an outcome in any admission data records:"
    section6_page1_2_16 = str(pat_num_dead) + " patients"
    section6_page1_2_17 = "Overall mortality:"
    section6_page1_2_18 = str(per_mortal)
    section6_page1_2 = [section6_page1_2_1, 
                        iden1_op + "<i>" + section6_page1_2_2 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section6_page1_2_3 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + bold_blue_ital_op + section6_page1_2_4 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden2_op + "<i>" + section6_page1_2_5 + "</i>" + iden_ed, 
                        iden2_op + "<i>" + bold_blue_ital_op + section6_page1_2_6 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden2_op + "<i>" + section6_page1_2_7 + "</i>" + iden_ed, 
                        iden2_op + "<i>" + bold_blue_ital_op + section6_page1_2_8 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        section6_page1_2_9, 
                        iden1_op + "<i>" + section6_page1_2_10 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section6_page1_2_11 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + bold_blue_ital_op + section6_page1_2_12 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section6_page1_2_13 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + bold_blue_ital_op + section6_page1_2_14 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden2_op + "<i>" + section6_page1_2_15 + "</i>" + iden_ed, 
                        iden2_op + "<i>" + bold_blue_ital_op + section6_page1_2_16 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden2_op + "<i>" + section6_page1_2_17 + "</i>" + iden_ed, 
                        iden2_op + "<i>" + bold_blue_ital_op + section6_page1_2_18 + bold_blue_ital_ed + "</i>" + iden_ed]
    section3_page2_org = ["<b>" + "Organism" + "</b>", 
                        "<b>" + "" + "</b>", 
                        "<b>" + "" + "</b>", 
                        "<b>" + "" + "</b>", 
                        "<b>" + s_aureus + "</b>", 
                        "<b>" + ent_spp + "</b>", 
                        "<b>" + s_pneumoniae + "</b>", 
                        "<b>" + sal_spp + "</b>", 
                        "<b>" + e_coli + "</b>", 
                        "<b>" + k_pneumoniae + "</b>", 
                        "<b>" + p_aeruginosa + "</b>", 
                        "<b>" + aci_spp + "</b>", 
                        "<b>" + "Total:" + "</b>"]
    ##Page3-8
    section6_page2_1_1 = "The AMASS application merged the microbiology data file and hospital admission data file. " + \
                        "The merged dataset was then de−duplicated so that only the first isolate per patient per specimen per reporting period was included in the analysis. " + \
                        "The de−duplicated data was stratified by infection origin (community−origin infection or hospital−origin infection)."
    section6_page2_1 = [section6_page2_1_1]
    section6_page2_2_1 = "The following figures and tables show the mortality of patients who were blood culture positive for antimicrobial non−susceptible and susceptible isolates."
    section6_page2_2 = [section6_page2_2_1]
    ##Page3
    section6_page3_1_1 = "*Mortality is the proportion (%) of in−hospital deaths (all−cause deaths). " + \
                        "This represents the number of in−hospital deaths (numerator) over the total number of patients with blood culture positive for the organism and the type of pathogen (denominator). " + \
                        "The AMASS application de−duplicates the data by included only the first isolate per patient per specimen type per evaluation period. " + \
                        "NS=non−susceptible; S=susceptible; CI=confidence interval"
    #DL                    
    section6_page3_1_2 = "Fluoroquinolone−NS=NS to any fluoroquinolone tested"
    section6_page3_1_3 = "**3GC-NS [for this section]: NS to any 3rd-generation cephalosporin excluding isolates which are non-susceptible to carbapenem. " + \
                        "***3GC-S [for this section]: S to all 3rd-generation cephalosporin tested excluding isolates which are non-susceptible to carbapenem."
    #CL
    # section6_page3_1_2 = "Penicillin−NS=NS to Penicillin G tested; Fluoroquinolone−NS=NS to any fluoroquinolone tested"
    # section6_page3_1_3 = "**3GC-NS [for this section]: NS to any 3rd-generation cephalosporin and NS to any carbapenem (patients without AST results for 3GC or carbapenem were not counted in the denominator). " + \
    #                     "***3GC-S [for this section]: S to any 3rd-generation cephalosporin tested."

    section6_page3_1_4 = "Carbapenem-NS=NS to any Carbapenem tested"
    section6_page3_1 = [section6_page3_1_1]
    section6_page3_2 = [section6_page3_1_1 + "; " + section6_page3_1_2]
    section6_page3_3 = [section6_page3_1_1 + "; " + section6_page3_1_3]
    section6_page3_4 = [section6_page3_1_1 + "; " + section6_page3_1_4]

    ######### SECTION6: PAGE1 #########
    report_title(c,'Section [6] Mortality involving AMR and',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'antimicrobial−susceptible infections',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,section6_page1_1, 1.0*inch, 6.0*inch, 460, 250,font_size=11)
    report_title(c,'Results',1.07*inch, 5.9*inch,'#3e4444',font_size=12)
    report_context(c,section6_page1_2, 1.0*inch, 0.7*inch, 460, 370,font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()
    ######### SECTION6: PAGE2 #########
    report_context(c,section6_page2_1, 1.0*inch, 8.5*inch, 460, 120,font_size=11)
    report_context(c,section3_page2_org, 1.0*inch, 5.0*inch, 460, 250, font_size=11, line_space=18.0)
    table_draw = report3_table(result_mor_table)
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 3.2*inch, 5.9*inch)
    report_context(c,section6_page2_2, 1.0*inch, 4.0*inch, 460, 50,font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.showPage()
    ######### SECTION6: PAGE3-6 #########
    d_head = {"com_1":[lst_org_format[0], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, df_numpat_com.loc[lst_org_full[0],"Total_number_of_patients"], 5.5*inch, 9.0*inch],
              "hos_1":[lst_org_format[0], 1.0*inch, 7.2*inch, 4.0*inch, 7.2*inch, df_numpat_hos.loc[lst_org_full[0],"Total_number_of_patients"], 5.5*inch, 7.2*inch],
              "com_2":[lst_org_format[1], 1.0*inch, 5.4*inch, 4.0*inch, 5.4*inch, df_numpat_com.loc[lst_org_full[1],"Total_number_of_patients"], 5.5*inch, 5.4*inch],
              "hos_2":[lst_org_format[1], 1.0*inch, 3.6*inch, 4.0*inch, 3.6*inch, df_numpat_hos.loc[lst_org_full[1],"Total_number_of_patients"], 5.5*inch, 3.6*inch],
              "com_3":[lst_org_format[2], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, df_numpat_com.loc[lst_org_full[2],"Total_number_of_patients"], 5.5*inch, 9.0*inch],
              "hos_3":[lst_org_format[2], 1.0*inch, 7.2*inch, 4.0*inch, 7.2*inch, df_numpat_hos.loc[lst_org_full[2],"Total_number_of_patients"], 5.5*inch, 7.2*inch],
              "com_4":[lst_org_format[3], 1.0*inch, 5.4*inch, 4.0*inch, 5.4*inch, df_numpat_com.loc[lst_org_full[3],"Total_number_of_patients"], 5.5*inch, 5.4*inch],
              "hos_4":[lst_org_format[3], 1.0*inch, 3.6*inch, 4.0*inch, 3.6*inch, df_numpat_hos.loc[lst_org_full[3],"Total_number_of_patients"], 5.5*inch, 3.6*inch],
              "com_5":[lst_org_format[4], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, df_numpat_com.loc[lst_org_full[4],"Total_number_of_patients"], 5.5*inch, 9.0*inch],
              "hos_5":[lst_org_format[4], 1.0*inch, 7.2*inch, 4.0*inch, 7.2*inch, df_numpat_hos.loc[lst_org_full[4],"Total_number_of_patients"], 5.5*inch, 7.2*inch],
              "com_6":[lst_org_format[5], 1.0*inch, 5.4*inch, 4.0*inch, 5.4*inch, df_numpat_com.loc[lst_org_full[5],"Total_number_of_patients"], 5.5*inch, 5.4*inch],
              "hos_6":[lst_org_format[5], 1.0*inch, 3.6*inch, 4.0*inch, 3.6*inch, df_numpat_hos.loc[lst_org_full[5],"Total_number_of_patients"], 5.5*inch, 3.6*inch],
              "com_7":[lst_org_format[6], 1.0*inch, 9.0*inch, 4.0*inch, 9.0*inch, df_numpat_com.loc[lst_org_full[6],"Total_number_of_patients"], 5.5*inch, 9.0*inch],
              "hos_7":[lst_org_format[6], 1.0*inch, 7.2*inch, 4.0*inch, 7.2*inch, df_numpat_hos.loc[lst_org_full[6],"Total_number_of_patients"], 5.5*inch, 7.2*inch],
              "com_8":[lst_org_format[7], 1.0*inch, 5.4*inch, 4.0*inch, 5.4*inch, df_numpat_com.loc[lst_org_full[7],"Total_number_of_patients"], 5.5*inch, 5.4*inch],
              "hos_8":[lst_org_format[7], 1.0*inch, 3.6*inch, 4.0*inch, 3.6*inch, df_numpat_hos.loc[lst_org_full[7],"Total_number_of_patients"], 5.5*inch, 3.6*inch]}
    d_pic =  {"com_1":[lst_org_short[0], 1.5*inch, 6.0*inch, 2.8*inch, 5.3*inch, "community"],
              "hos_1":[lst_org_short[0], 1.5*inch, 4.2*inch, 2.8*inch, 5.3*inch, "hospital"],
              "com_2":[lst_org_short[1], 0.8*inch, 2.3*inch, 3.5*inch, 5.5*inch, "community"],
              "hos_2":[lst_org_short[1], 0.8*inch, 0.4*inch, 3.5*inch, 5.5*inch, "hospital"],
              "com_3":[lst_org_short[2], 0.9*inch, 6.0*inch, 3.4*inch, 5.3*inch, "community"],
              "hos_3":[lst_org_short[2], 0.9*inch, 4.2*inch, 3.4*inch, 5.3*inch, "hospital"],
              "com_4":[lst_org_short[3], 0.4*inch, 2.3*inch, 3.9*inch, 5.5*inch, "community"],
              "hos_4":[lst_org_short[3], 0.4*inch, 0.4*inch, 3.9*inch, 5.5*inch, "hospital"],
              "com_5":[lst_org_short[4], 0.9*inch, 6.0*inch, 3.4*inch, 5.3*inch, "community"],
              "hos_5":[lst_org_short[4], 0.9*inch, 4.2*inch, 3.4*inch, 5.3*inch, "hospital"],
              "com_6":[lst_org_short[5], 0.9*inch, 2.3*inch, 3.4*inch, 5.5*inch, "community"],
              "hos_6":[lst_org_short[5], 0.9*inch, 0.45*inch, 3.4*inch, 5.5*inch, "hospital"],
              "com_7":[lst_org_short[6], 0.8*inch, 6.0*inch, 3.5*inch, 5.3*inch, "community"],
              "hos_7":[lst_org_short[6], 0.8*inch, 4.2*inch, 3.5*inch, 5.3*inch, "hospital"],
              "com_8":[lst_org_short[7], 0.8*inch, 2.3*inch, 3.5*inch, 5.5*inch, "community"],
              "hos_8":[lst_org_short[7], 0.8*inch, 0.4*inch, 3.5*inch, 5.5*inch, "hospital"]}
    d_table = {"com_1":[org1_com, 500, 300, 4.7*inch, 8.4*inch],
               "hos_1":[org1_hos, 500, 300, 4.7*inch, 6.6*inch],
               "com_2":[org2_com, 500, 300, 4.7*inch, 4.8*inch],
               "hos_2":[org2_hos, 500, 300, 4.7*inch, 2.9*inch],
               
               "com_3":[org3_com, 500, 300, 4.7*inch, 8.4*inch],
               "hos_3":[org3_hos, 500, 300, 4.7*inch, 6.6*inch],
               "com_4":[org4_com, 500, 300, 4.7*inch, 4.8*inch],
               "hos_4":[org4_hos, 500, 300, 4.7*inch, 2.9*inch],
               
               "com_5":[org5_com, 500, 300, 4.7*inch, 8.3*inch],
               "hos_5":[org5_hos, 500, 300, 4.7*inch, 6.5*inch],
               "com_6":[org6_com, 500, 300, 4.7*inch, 4.7*inch],
               "hos_6":[org6_hos, 500, 300, 4.7*inch, 2.9*inch],
               
               "com_7":[org7_com, 500, 300, 4.7*inch, 8.4*inch],
               "hos_7":[org7_hos, 500, 300, 4.7*inch, 6.6*inch],
               "com_8":[org8_com, 500, 300, 4.7*inch, 4.8*inch],
               "hos_8":[org8_hos, 500, 300, 4.7*inch, 2.9*inch]}
    d_note = {"com_1":["no_note"],
               "hos_1":["no_note"],
               "com_2":["no_note"],
               "hos_2":[section6_page3_1, 1.0*inch, 0.4*inch, 460, 130, lst_pagenumber[2], lastpage],
               "com_3":["no_note"],
               "hos_3":["no_note"],
               "com_4":["no_note"],
               "hos_4":[section6_page3_2, 1.0*inch, 0.4*inch, 460, 130, lst_pagenumber[3], lastpage],
               "com_5":["no_note"],
               "hos_5":["no_note"],
               "com_6":["no_note"],
               "hos_6":[section6_page3_3, 1.0*inch, 0.4*inch, 460, 130, lst_pagenumber[4], lastpage],
               "com_7":["no_note"],
               "hos_7":["no_note"],
               "com_8":["no_note"],
               "hos_8":[section6_page3_4, 1.0*inch, 0.4*inch, 460, 130, lst_pagenumber[5], lastpage]}
    for keys in d_head.keys():
        report_title(c,'Section [6] Mortality involving AMR and',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
        report_title(c,'antimicrobial−susceptible infections',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
        line_1_1 = ["<b>" + "Blood: " + d_head[keys][0] + "</b>"]
        line_1_2 = ["<b>" + d_pic[keys][5].capitalize() + "-origin" + "</b>"]
        line_1_3 = [bold_blue_ital_op + "( No. of patients = " + str(d_head[keys][5]) + " )" + bold_blue_ital_ed]
        report_context(c,line_1_1, d_head[keys][1], d_head[keys][2], 300, 50, font_size=12, font_align=TA_LEFT)
        report_context(c,line_1_2, d_head[keys][3], d_head[keys][4], 200, 50, font_size=12, font_align=TA_LEFT)
        report_context(c,line_1_3, d_head[keys][6], d_head[keys][7], 200, 50, font_size=12, font_align=TA_LEFT)
        c.drawImage(path_result+'Report6_mortality_'+ d_pic[keys][0] + "_" + d_pic[keys][5]+".png", d_pic[keys][1], d_pic[keys][2], preserveAspectRatio=True, width=d_pic[keys][3], height=d_pic[keys][4],showBoundary=False) 
        table_draw = report2_table_nons(d_table[keys][0])
        table_draw.wrapOn(c, d_table[keys][1], d_table[keys][2])
        table_draw.drawOn(c, d_table[keys][3], d_table[keys][4])
        if d_note[keys][0] != "no_note":
            report_context(c,d_note[keys][0], d_note[keys][1], d_note[keys][2], d_note[keys][3], d_note[keys][4], font_size=9,line_space=12)
            report_todaypage(c,55,30,"Created on: "+today)
            report_todaypage(c,270,30,"Page " + d_note[keys][5] + " of " + d_note[keys][6])
            c.showPage()
        else:
            pass

def annexA_page1to2(result_table, org_table, pat_table,
                    lst_pagenumber=pagenumber_ava_annexA, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden3_op = "<para leftindent=\"105\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    tab1st = "&nbsp;"
    tab4th = "&nbsp;&nbsp;&nbsp;&nbsp;"
    ##variables
    spc_date_start  = result_table.loc[result_table["Parameters"]=="Minimum_date","Values"].tolist()[0]
    spc_date_end    = result_table.loc[result_table["Parameters"]=="Maximum_date","Values"].tolist()[0]
    spc_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_all_culture_positive","Values"].tolist()[0]
    blo_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_positive","Values"].tolist()[0]
    csf_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_csf_culture_positive","Values"].tolist()[0]
    gen_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_genital_swab_culture_positive","Values"].tolist()[0]
    res_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_rts_culture_positive","Values"].tolist()[0]
    sto_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_stool_culture_positive","Values"].tolist()[0]
    uri_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_urine_culture_positive","Values"].tolist()[0]
    oth_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_others_culture_positive","Values"].tolist()[0]
    
    ##Page1
    annexA_page1_1_1 = "This supplementary report has two parts; including (A1) isolate-based notifiable bacterial infections and (A2) mortality involving notifiable bacterial infections. The isolate-based notifiable bacterial infections supplementary report is generated by default, even if the hospital_admission_data file is unavailable. This is to enable hospitals with only microbiology data available to utilize the de-duplication and report generation functions of AMASS."
    annexA_page1_1_2 = "Please note that the completion of this supplementary report is strongly associated with the availability of data (particularly, all bacterial pathogens and all types of specimens) and the completion of the data dictionary files to make sure that the AMASS application understands the notifiable bacteria and each type of specimens."
    annexA_page1_1_3 = "Annex A includes various type of specimens including blood, cerebrospinal fluid (CSF), respiratory tract specimens, urine, genital swab, stool and other or unknown sample types. The microorganisms in this report were initially selected from common notifiable bacterial diseases in Thailand."
    annexA_page1_1 = [annexA_page1_1_1, 
                    add_blankline + annexA_page1_1_2, 
                    add_blankline + annexA_page1_1_3]

    annexA_page1_2_1 = "Note: The list of notifiable bacteria included in the AMASS application version 2.0 was generated based on the literature review and the collaboration with Department of Disease Control, Ministry of Public Health, Thailand. The list could be expanded or modified in the next version of AMASS."
    annexA_page1_2 = [green_op + annexA_page1_2_1 + green_ed]
    ##Page2
    annexA_page2_1_1 = "The microbiology_data file had:"
    annexA_page2_1_2 = "Sample collection dates ranged from " + \
                        bold_blue_ital_op + str(spc_date_start) + bold_blue_ital_ed + \
                        "  to  " + \
                        bold_blue_ital_op + str(spc_date_end) + bold_blue_ital_ed
    annexA_page2_1_3 = "Number of records of clinical specimens collected with culture positive for a notifiable bacteria under this survey:"
    annexA_page2_1_4 = bold_blue_ital_op + str(spc_num_pos_plus) + bold_blue_ital_ed + "  specimen records (" + \
                        bold_blue_ital_op + str(blo_num_pos_plus) + " , " + str(csf_num_pos_plus) + " , " + \
                        str(gen_num_pos_plus) + " , " + str(res_num_pos_plus) + " , " + str(sto_num_pos_plus) + " , " + \
                        str(uri_num_pos_plus) + " , " + str(oth_num_pos_plus) + bold_blue_ital_ed + \
                        " were blood, CSF, genital swab, respiratory tract specimens, stool, urine, and other or unknown sample types, respectively) "
    annexA_page2_1_5 = "The AMASS application de-duplicated the data by including only the first isolate per patient per specimen type per evaluation period as described in the method. The number of patients with positive samples is as follows:"
    annexA_page2_1 = [annexA_page2_1_1, 
                    iden1_op + "<i>" + annexA_page2_1_2 + "</i>" + iden_ed, 
                    iden1_op + "<i>" + annexA_page2_1_3 + "</i>" + iden_ed, 
                    iden2_op + "<i>" + annexA_page2_1_4 + "</i>" + iden_ed, 
                    add_blankline + annexA_page2_1_5]
    annexA_page2_2_1 = "*Some patients may have more than one type of clinical specimen culture positive for the notifiable bacteria under the survey, and some may have more than one notifiable organism per evaluation period."
    annexA_page2_2_2 = "CSF = Cerebrospinal fluid; RTS = Respiratory tract specimens; Others = Other or unknown sample types;"
    annexA_page2_2_3 = "NA = Not applicable (i.e. the specimen type is not available or identified in the microbiology_data file)"
    annexA_page2_2 = [annexA_page2_2_1,
                    annexA_page2_2_2, 
                    annexA_page2_2_3]
    ######### ANNEX A: PAGE1 ##########
    report_title(c,'Annex A: Supplementary report on notifiable bacterial',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'infections',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,annexA_page1_1, 1.0*inch, 4.5*inch, 460, 350,font_size=11)
    report_title(c,'Notifiable bacteria under the survey',1.07*inch, 4.8*inch,'#3e4444',font_size=12)
    table_draw = report_table_annexA_page1(org_table)
    table_draw.wrapOn(c, 700, 700)
    table_draw.drawOn(c, 1.2*inch, 3.0*inch)
    report_context(c,annexA_page1_2, 1.0*inch, 0.5*inch, 460, 100,font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()
    ######### ANNEX A: PAGE2 ##########
    report_title(c,'Annex A1: Isolated-based notifiable bacterial infections',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'Results',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,annexA_page2_1, 1.0*inch, 5.9*inch, 460, 250,font_size=11)
    table_draw = report_table_annexA_page2(pat_table)
    table_draw.wrapOn(c, 400, 300)
    if len(pat_table) < 7:
        table_draw.drawOn(c, 1.2*inch, 3.0*inch)
    else:
        table_draw.drawOn(c, 1.2*inch, 2.5*inch)
    report_context(c,annexA_page2_2, 1.0*inch, 1.0*inch, 460, 70,font_size=9,line_space=12)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.showPage()

def annexA_page3(mor_table, lst_pagenumber=pagenumber_ava_annexA, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden3_op = "<para leftindent=\"105\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    tab1st = "&nbsp;"
    tab4th = "&nbsp;&nbsp;&nbsp;&nbsp;"
    ##variables
    ##Page3
    annexA_page3_1_1 = "A report on mortality involving notifiable bacterial infections is generated only if data on patient outcomes (i.e. discharge status) are available. The term \"mortality involving notifiable bacterial infections\" was used because the mortality reported was all-cause mortality. This measure of mortality included deaths caused by or related to other underlying and intermediate causes. The AMASS application merged the microbiology data file and hospital admission data file. The merged dataset was then de-duplicated so that only the first isolate per patient per specimen per reporting period was included in the analysis."
    annexA_page3_1 = [annexA_page3_1_1]
    annexA_page3_2_1 = "*Mortality is the proportion (%) of in-hospital deaths (all-cause deaths). This represents the number of in-hospital deaths (numerator) over the total number of patients with culture positive for each type of pathogen (denominator). Some patients may have the data of a clinical specimen culture positive for the notifiable bacteria under the survey in the microbiology data file, but do not have the data in the hospital admission data file. That is the most common cause of the discrepancy between total number of patients with notifiable bacterial infections presented in the Annex A1 and the Annex A2 (followed by typos in patient identifiers in either data file)."
    annexA_page3_2_2 = "CI = confidence interval"
    annexA_page3_2 = [annexA_page3_2_1, 
                    annexA_page3_2_2]
    ######### ANNEX A: PAGE3 ##########
    report_title(c,'Annex A2: Mortality involving notifiable bacterial infections',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_context(c,annexA_page3_1, 1.0*inch, 7.5*inch, 460, 150,font_size=11)
    report_context(c,["*Mortality (%)"], 2.0*inch, 2.2*inch, 150, 30, font_size=9, font_align=TA_CENTER, line_space=14)
    c.drawImage(path_result+"AnnexA_mortality.png", 1.2*inch, 2.5*inch, preserveAspectRatio=False, width=2.5*inch, height=5.0*inch,showBoundary=False) 
    table_draw = report_table_annexA_page3(mor_table)
    table_draw.wrapOn(c, 265, 300)
    if len(mor_table) < 7:
        table_draw.drawOn(c, 4.2*inch, 5.0*inch)
    else:
        table_draw.drawOn(c, 4.2*inch, 3.5*inch)
    report_context(c,annexA_page3_2, 1.0*inch, 0.7*inch, 460, 100,font_size=9,line_space=12)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[2] + " of " + lastpage)
    c.showPage()

def annexB(blo_table, blo_table_bymonth, lst_pagenumber=pagenumber_ava_annexB, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden3_op = "<para leftindent=\"105\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    tab1st = "&nbsp;"
    tab4th = "&nbsp;&nbsp;&nbsp;&nbsp;"
    ##variables
    ##Page1
    annexB_page1_1_1 = "This supplementary report is generated by default, even if the hospital_admission_data file is unavailable. The management of clinical and laboratory practice can be supported by some data indictors such as blood culture contamination rate, proportion of notifiable antibiotic-pathogen combinations, and proportion of isolates with infrequent phenotypes or potential errors in AST results. Isolates with infrequent phenotypes or potential errors in AST results include (a) reports of organisms which are intrinsically resistant to an antibiotic but are reported as susceptible and (b) reports of organisms with discordant AST results. "
    annexB_page1_1_2 = "This supplementary report could support the clinicians, policy makers and the laboratory staff to understand their summary data quickly. The laboratory staff could also use \"Supplementary_data_indicators_report.pdf\" generated in the folder \"Report_with_patient_identifiers\" to check and validate individual data records further. "
    annexB_page1_1_3 = "<b>This supplementary report was estimated from data of blood specimens only.</b> Please note that the data indicators do not represent quality of the clinical or laboratory practice."
    annexB_page1_1 = [annexB_page1_1_1, 
                    add_blankline + annexB_page1_1_2, 
                    add_blankline + annexB_page1_1_3]

    annexB_page1_2_1 = "*Blood culture contamination rate is defined as the number of raw contaminated cultures per number of blood cultures received by the laboratory per reporting period. Blood culture contamination rate will not be estimated in case that the data of negative culture (specified as 'no growth' in the dictionary_for_microbiology_data file) is not available. "
    annexB_page1_2_2 = "**Notifiable antibiotic-pathogen combinations and their classifications are defined as WHO list of AMR priority pathogen published in 2017. "
    annexB_page1_2_3 = "**, ***The proportion is estimated per number of blood specimens culture positive for any organisms with AST result in the raw microbiology data. "
    annexB_page1_2_4 = "*, **, ***Details of the criteria are available in Table 3 and Table 4 of \"Supplementary_data_indicators_report.pdf\", and \"list_of_indicators.xlsx\" in the folder \"Configuration\". "
    annexB_page1_2_5 = "NA = Not applicable"
    annexB_page1_2 = [annexB_page1_2_1+annexB_page1_2_2+annexB_page1_2_3+annexB_page1_2_4+annexB_page1_2_5]
    ##Page2
    annexB_page2_1_2 = "Data was stratified by month to assist detection of missing data and understand the change of indicators by months."
    annexB_page2_1 = [annexB_page2_1_2]
    ########### ANNEX B: PAGE1 ########
    report_title(c,"Annex B: Supplementary report on data indicators",1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,annexB_page1_1, 1.0*inch, 4.5*inch, 460, 350, font_size=11)
    report_title(c,'Results',1.07*inch, 5.2*inch,'#3e4444',font_size=12)
    table_draw = report_table_annexB_page1(blo_table)
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.07*inch, 2.5*inch)
    report_context(c,annexB_page1_2, 1.0*inch, 0.6*inch, 460, 130, font_size=9, line_space=14)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()
    ########### ANNEX B: PAGE2 ########
    report_title(c,"Annex B: Supplementary report on data indicators",1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'Reporting period by months',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    report_context(c,annexB_page2_1, 1.0*inch, 8.7*inch, 460, 50, font_size=11)
    table_draw = report_table_annexB(blo_table_bymonth)
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.2*inch, 4.8*inch)
    report_context(c,annexB_page1_2, 1.0*inch, 0.8*inch, 460, 130, font_size=9,line_space=14)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.showPage()

def method(lst_pagenumber=pagenumber_ava_other, lst_org=lst_org_format, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden3_op = "<para leftindent=\"105\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    tab1st = "&nbsp;"
    tab4th = "&nbsp;&nbsp;&nbsp;&nbsp;"
    ##variables
    ##Page1
    method_page1_1 = "<b>" + "Data source:" + "</b>"
    method_page1_2 = "For each run (double−click on AMASS.bat file), the AMASS application used the microbiology data file (microbiology_data) and the hospital admission data file (hospital_admission_data) that were stored in the same folder as the application file. " + \
                    "Hence, if the user would like to update, correct, revise or change the data, the data files in the folder should be updated before the AMASS.bat file is double−clicked again. " + \
                    "A new report based on the updated data would then be generated."
    method_page1_3 = add_blankline + "<b>" + "Requirements:" + "</b>"
    method_page1_4 = "<b>" + "− Computer with Microsoft Windows 7 or 10" + "</b>"
    method_page1_5 = "AMASS may work in other versions of Microsoft Windows and other operating systems. " + \
                    "However, thorough testing and adjustment have not been performed."
    method_page1_6 = "<b>" + "− AMASSv2.0.zip package file" + "</b>"
    method_page1_7 = "The AMASS application is to be downloaded from " + "<u><link href=\"https://www.amass.website\" color=\"blue\"fontName=\"Helvetica\">https://www.amass.website</link></u>" + ", and unzipped to generate an AMASS folder that could be stored under any folder in the computer. " + \
                    "The AMASS folder contains 3 files (AMASS.bat, dictionary_for_microbiology_data.xlsx, and dictionary_for_hospital_admission_data.xlsx), and 5 folders (Configuration, Example_Dataset_1_WHONET, Example_Dataset_2, Example_Dataset_3_longformat, Programs)."
    method_page1_8 = "<b>" + "− Microbiology data file (microbiology_data in .csv or .xlsx file format)" + "</b>"
    method_page1_9 = "The user needs to obtain microbiology data, and then copy & paste this data file into the same folder as the AMASS.bat file."
    method_page1_10 = "<b>" + "− [Optional] Hospital admission data file (hospital_admission_data)" + "</b>"
    method_page1_11 = "If available, the user could obtain hospital admission data, and then copy & paste this data file into the same folder as the AMASS.bat file."
    method_page1_12 = add_blankline + "<b>" + "Not required:" + "</b>"
    method_page1_13 = "<b>" + "− Internet to run AMASS application" + "</b>"
    method_page1_14 = "The AMASS application will run offline. No data are transferred while the application is running and reports are being generated; the reports are in PDF format (do not contain any patient identifier) and can be shared under the user's jurisdiction."
    method_page1_15 = "<b>" + "− R and Python" + "</b>"
    method_page1_16 = "The download package (AMASSv2.0.zip) included R portable, Python portable and their libraries that the AMASS application requires. " + \
                    "The user does not need to install any programme before using the AMASS. " + \
                    "The user also does not have to uninstall R or Python if the computer already has the programme installed. " + \
                    "The user does not need to know how to use R and Python."
    method_page1 = [method_page1_1, method_page1_2, method_page1_3, method_page1_4, method_page1_5, 
                    method_page1_6, method_page1_7, method_page1_8, method_page1_9, method_page1_10, 
                    method_page1_11, method_page1_12, method_page1_13, method_page1_14, method_page1_15, 
                    method_page1_16]
    ##Page2
    method_page2_1_1 = green_op + "<b>" + "Note:" + "</b>" + green_ed
    method_page2_1_2 = green_op + "[1] Please ensure that the file names of microbiology data file (microbiology_data) and the hospital admission data file (hospital_admission_data) are identical to what is written here. " + \
                    "Please make sure that all are lower−cases with an underscore '_' at each space." + green_ed
    method_page2_1_3 = green_op + "[2] Please ensure that both microbiology and hospital admission data files have no empty rows before the row of the variable names (i.e. the variable names are the first row in both files)." + green_ed
    method_page2_1_4 = green_op + "[3] For the first run, an user may need to fill the data dictionary files to make sure that the AMASS application understands your variable names and values." + green_ed
    method_page2_1 = [method_page2_1_1, method_page2_1_2, method_page2_1_3, method_page2_1_4]
    method_page2_2_1 = "AMASS uses a tier−based approach. In cases when only the microbiology data file with the results of culture positive samples is available, only section one and two would be generated for users. " + \
                    "Section three would be generated only when data on admission date are available. " + \
                    "This is because these data are required for the stratification by origin of infection. " + \
                    "Section four would be generated only when data of specimens with culture negative (no microbial growth) are available in the microbiology data. " + \
                    "This is because these data are required for the sample−based approach. " + \
                    "Section five would be generated only when both data of specimens with culture negative and admission date are available. " + \
                    "Section six would be generated only when mortality data are available."
    method_page2_2_2 = add_blankline + "Mortality was calculated from the number of in−hospital deaths (numerator) over the total number of patients with blood culture positive for the organism (denominator). " + \
                    "Please note that this is the all−cause mortality calculated using the outcome data in the data file, and may not necessarily represent the mortality directly due to the infections."
    method_page2_2 = [method_page2_2_1, method_page2_2_2]
    method_page2_3_1 = "<b>" + "How to use data dictionary files" + "</b>"
    method_page2_3_2 = "In cases when variable names in the microbiology and hospital admission data files were not the same as the one that AMASS used, the data dictionary files could be edited. " + \
                    "The raw microbiology and hospital admission data files were to be left unchanged. " + \
                    "The data dictionary files provided could be edited and re−used automatically when the microbiology and hospital admission data files were updated and the AMASS.bat were to be double−clicked again (i.e. the data dictionary files would allow the user to re−analyze data files without the need to adjust variable names and data value again every time)."
    method_page2_3 = [method_page2_3_1, method_page2_3_2]
    ##Page3
    method_page3_1_1 = "For example:"
    method_page3_1_2 = "If variable name for 'hospital number' is written as 'hn' in the raw data file, the user would need to add 'hn' in the cell next to 'hospital_number'. " + \
                    "If data value for blood specimens is defined by 'Blood−Hemoculture' in the raw data file, then the user would need to add 'Blood−Hemoculture' in the cell next to 'blood_specimen'."
    method_page3_1 = [method_page3_1_1, method_page3_1_2]
    method_page3_2 = ["<b>" + "Dictionary file (dictionary_for_microbiology_data.xlsx) may show up as in the table below:" + "</b>"]
    table_med_1 = [["Variable names used in AMASS", "Variable names used in \n your microbiology data file", "Requirements"],
                ["Don't change values in this \n column, but you can add rows \n with similar values if you need", 
                    "Change values in this column to \n represent how variable names \n are written in your raw \n microbiology data file", ""], 
                ["hospital_number", "", "Required"], 
                ["Values described in AMASS", "Values used in your \n microbiology data file", "Requirements"], 
                ["blood_specimen", "", "Required"]]

    method_page3_3 = ["<b>" + "Please fill in your variable names as follows:" + "</b>"]
    table_med_2 = [["Variable names used in AMASS", "Variable names used in \n your microbiology data file", "Requirements"],
                ["Don't change values in this \n column, but you can add rows \n with similar values if you need", 
                    "Change values in this column to \n represent how variable names \n are written in your raw \n microbiology data file", ""], 
                ["hospital_number", "hn", "Required"], 
                ["Values described in AMASS", "Values used in your \n microbiology data file", "Requirements"], 
                ["blood_specimen", "Blood−Hemoculture", "Required"]]

    method_page3_4 = ["Then, save the file. For every time the user double−clicked AMASS.bat, the application would know that the variable named 'hn' is similar to 'hospital_number' and represents the patient identifier in the analysis."]
    ##Page4
    method_page4_1 = ["<b>" + "Organisms included for the AMR Surveillance Report:" + "</b>", 
                    "− " + lst_org[0], 
                    "− " + lst_org[1], 
                    "− " + lst_org[2], 
                    "− " + lst_org[3], 
                    "The eight organisms and antibiotics included in the report were selected based on the global priority list of antibiotic resistant bacteria and Global Antimicrobial Resistance Surveillance System (GLASS) of WHO [1,2]."]
    method_page4_2 = ["− " + lst_org[4], 
                    "− " + lst_org[5], 
                    "− " + lst_org[6], 
                    "− " + lst_org[7]]
    method_page4_5_1 = "<b>" + "Definitions:" + "</b>"
    method_page4_5_2 = "The definitions of infection origin proposed by the WHO GLASS was used [1]. In brief, community−origin bloodstream infection (BSI) was defined for patients in the hospital within the first two calendar days of admission when the first blood culture positive specimens were taken. " + \
                    "Hospital−origin BSI was defined for patients in the hospital longer than the first two calendar days of admission when the first blood culture positive specimens were taken. " + \
                    "In cases when the user had additional data on infection origin defined by infection control team or based on referral data, the user could edit the data dictionary file (variable name \'infection_origin\') and the AMASS application would use the data of that variable to stratify the data by origin of infection instead of the above definition. " + \
                    "However, in cases when data on infection origin were not available (as in many hospitals in LMICs), the above definition would be calculated based on admission date and specimen collection date (with cutoff of 2 calendar days) and used to classify infections as community−origin or hospital−origin."
    method_page4_5_3 = "<b>" + "De−duplication:" + "</b>"
    method_page4_5_4 = "When more than one blood culture was collected during patient management, duplicated findings of the same patient were excluded (de−duplicated). " + \
                    "Only one result was reported for each patient per sample type (blood) and surveyed organisms (listed above)." + \
                    "For example, if two blood cultures from the same patient had <i>E. coli</i>, only the first would be included in the report. " + \
                    "If there was growth of <i>E. coli</i> in one blood culture and of <i>K. pneumoniae</i> in the other blood culture, then both results would be reported. " + \
                    "One would be for the report on <i>E. coli</i> and the other one would be for the report on <i>K. pneumoniae</i>."
    method_page4_5 = [method_page4_5_1, method_page4_5_2, add_blankline + method_page4_5_3, method_page4_5_4]
    ##Backcover
    backcover_1_1 = "<b>" + "References:" + "</b>"
    backcover_1_2 = "[1] World Health Organization (2018) Global Antimicrobial Resistance Surveillance System (GLASS) Report. Early implantation 2016−2017. http://apps.who.int/iris/bitstream/handle/10665/259744/9789241513449−eng.pdf. (accessed on 3 Dec 2018)"
    backcover_1_3 = "[2] World Health Organization (2017) Global priority list of antibiotic−resistant bacteria to guide research, discovery, and development of new antibiotics. https://www.who.int/medicines/publications/WHO−PPL−Short_ Summary_25Feb−ET_NM_WHO.pdf. (accessed on 3 Dec 2018)"
    backcover_1 = [backcover_1_1, backcover_1_2, backcover_1_3]
    ########## METHOD: PAGE1 ##########
    report_title(c,'Methods used by the AMASS application',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_context(c,method_page1, 1.0*inch, 0.7*inch, 460, 680, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()
    ########## METHOD: PAGE2 ##########
    report_context(c,method_page2_1, 1.0*inch, 8.0*inch, 460, 200, font_size=11)
    report_context(c,method_page2_2, 1.0*inch, 4.0*inch, 460, 300, font_size=11)
    report_context(c,method_page2_3, 1.0*inch, 1.5*inch, 460, 200, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[1] + " of " + lastpage)
    c.showPage()
    ########## METHOD: PAGE3 ##########
    report_context(c,method_page3_1, 1.0*inch, 8.7*inch, 460, 120, font_size=11)
    report_context(c,method_page3_2, 1.0*inch, 8.0*inch, 460, 50, font_size=11)
    table_draw = Table(table_med_1,  style=[('FONT',(0,0),(-1,-1),'Helvetica'),
                                            ('FONT',(0,0),(2,0),'Helvetica-Bold'),
                                            ('FONT',(0,3),(-1,-2),'Helvetica-Bold'),
                                            ('FONTSIZE',(0,0),(-1,-1),11),
                                            ('TEXTCOLOR',(0,2),(-3,-3),colors.red),
                                            ('TEXTCOLOR',(0,4),(-3,-1),colors.red),
                                            ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                                            ('ALIGN',(0,0),(-1,-1),'CENTER'),
                                            ('VALIGN',(0,0),(-1,-1),'MIDDLE')])
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.0*inch, 6.0*inch)
    report_context(c,method_page3_3, 1.0*inch, 5.0*inch, 460, 50, font_size=11)
    table_draw = Table(table_med_2,  style=[('FONT',(0,0),(-1,-1),'Helvetica'),
                                            ('FONT',(0,0),(2,0),'Helvetica-Bold'),
                                            ('FONT',(0,3),(-1,-2),'Helvetica-Bold'),
                                            ('FONTSIZE',(0,0),(-1,-1),11),
                                            ('TEXTCOLOR',(0,2),(-3,-3),colors.red),
                                            ('TEXTCOLOR',(0,4),(-3,-1),colors.red),
                                            ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                                            ('ALIGN',(0,0),(-1,-1),'CENTER'),
                                            ('VALIGN',(0,0),(-1,-1),'MIDDLE')])
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.0*inch, 3.0*inch)
    report_context(c,method_page3_4, 1.0*inch, 1.7*inch, 460, 80, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[2] + " of " + lastpage)
    c.showPage()
    ########## METHOD: PAGE4 ##########
    report_context(c,method_page4_1, 1.0*inch, 9.0*inch, 460, 120, font_size=11)
    report_context(c,method_page4_2, 4.0*inch, 8.75*inch, 460, 120, font_size=11)
    report_context(c,method_page4_5, 1.0*inch, 3.0*inch, 460, 450, font_size=11)
    u = inch/10.0
    c.setLineWidth(2)
    c.setStrokeColor(black)
    p = c.beginPath()
    p.moveTo(70,220) # start point (x,y)
    p.lineTo(7.45*inch,220) # end point (x,y)
    c.drawPath(p, stroke=1, fill=1)
    report_context(c,backcover_1, 1.0*inch, 1.0*inch, 460, 150, font_size=9)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[3] + " of " + lastpage)
    c.showPage()

def investor(lst_pagenumber=pagenumber_ava_other, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden3_op = "<para leftindent=\"105\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    tab1st = "&nbsp;"
    tab4th = "&nbsp;&nbsp;&nbsp;&nbsp;"
    ##variables
    ##Investor
    invest_1_1 = "The AMASS application is being developed by Cherry Lim, Clare Ling, Elizabeth Ashley, Paul Turner, Rahul Batra, Rogier van Doorn, Soawapak Hinjoy, Sopon Iamsirithaworn, Susanna Dunachie, Tri Wangrangsimakul, Viriya Hantrakun, William Schilling, John Stelling, Jonathan Edgeworth, Guy Thwaites, Nicholas PJ Day, Ben Cooper and Direk Limmathurotskul."
    invest_1_2 = "AMASS version 2.0 is being developed by Chalida Rangsiwutisak, Cherry Lim, Paul Tuner, John Stelling and Direk Limmathurotsakul."
    invest_1_3 = "AMASS version 1.0 was funded by the Wellcome Trust (grant no. 206736 and 101103). C.L. is funded by a Research Training Fellowship (grant no. 206736) and D.L. is funded by an Intermediate Training Fellowship (grant no. 101103) from the Wellcome Trust."
    invest_1_4 = "AMASS version 2.0 was funded by the Wellcome Trust Institutional Translational Partnership Award- MORU"
    invest_1 = [invest_1_1, add_blankline+invest_1_2, add_blankline+invest_1_3, add_blankline+invest_1_4]
    
    invest_2_1 = "If you have any queries about AMASS, please contact:"
    invest_2_2 = "For technical information:"
    invest_2_3 = "Chalida Rangsiwutisak (chalida@tropmedes.ac),"
    invest_2_4 = "Cherry Lim (cherry@tropmedres.ac), and"
    invest_2_5 = "Direk Limmathurotsakul (direk@tropmedres.ac)"
    invest_2_6 = "For implementation of AMASS at your hospitals in Thailand:"
    invest_2_7 = "Preeyarach Klaytong (preeyarach@tropmedres.ac)"
    invest_2 = ["<b>"+invest_2_1+"</b>", 
                "<b>"+invest_2_2+"</b>", 
                invest_2_3, 
                invest_2_4, 
                invest_2_5, 
                "<b>"+add_blankline+invest_2_6+"</b>", 
                invest_2_7]
    ############# INVESTOR ############
    report_title(c,'Investigator team',1.07*inch, 9.0*inch,'#3e4444',font_size=16)
    report_context(c,invest_1, 1.0*inch, 4.5*inch, 460, 300, font_size=11)
    report_context(c,invest_2, 1.0*inch, 1.5*inch, 460, 200, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[4] + " of " + lastpage)
    c.rotate(90)
    c.showPage()
    c.showPage()

def section1to3_nomicro(lst_pagenumber_1=pagenumber_ava_1, lst_pagenumber_2=pagenumber_ava_2, 
                        lst_pagenumber_3=pagenumber_ava_3, lastpage="47", today=date.today().strftime("%d %b %Y")):
    page1_1_1 = "Not applicable because microbiology_data.xlsx file is not available or the format of microbiology_data file is not supported. Please save microbiology_data file in excel format (.xlsx) or csv (.csv; UTF-8)."
    page1_1 = [page1_1_1]
    ######### SECTION1: PAGE1 #########
    report_title(c,'Section [1]: Data overview',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_context(c,page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber_1[0] + " of " + lastpage)
    c.showPage()
    ######### SECTION2: PAGE1 #########
    report_title(c,'Section [2]: Isolate−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_context(c,page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber_2[0] + " of " + lastpage)
    c.showPage()
    ######### SECTION3: PAGE1 #########
    report_title(c,'Section [3]: Isolate−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    report_context(c,page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: " + today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber_3[0] + " of " + lastpage)
    c.showPage()

def section3_nohospital(lst_pagenumber=pagenumber_ava_3, lastpage="47", today=date.today().strftime("%d %b %Y")):
    section3_page1_1_1 = "Proportions of antimicrobial−resistance infection stratified by origin of infection is not calculated because hospital admission date data is not available and infection origin variable is not available."
    section3_page1_1 = [section3_page1_1_1] 
    ######### SECTION3: PAGE1 #########
    report_title(c,'Section [3]: Isolate−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    report_context(c,section3_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()

def section4_nohospital(lst_pagenumber=pagenumber_ava_4, lastpage="47", today=date.today().strftime("%d %b %Y")):
    section4_page1_1_1 = "Incidence of infections per 100,000 tested population is not calculated because data on blood specimen with no growth is not available."
    section4_page1_1 = [section4_page1_1_1]
    ######### SECTION4: PAGE1 #########
    report_title(c,'Report [4]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_context(c,section4_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()

def section5_nohospital(lst_pagenumber=pagenumber_ava_5, lastpage="47", today=date.today().strftime("%d %b %Y")):
    section5_page1_1_1 = "Incidence of infections per 100,000 tested population stratified by infection origin is not calculated because data on blood specimen with no growth is not available, or stratification by origin of infection cannot be done (due to hospital admission date variable is not available)."
    section5_page1_1 = [section5_page1_1_1]
    ######### SECTION5: PAGE1 #########
    report_title(c,'Report [5]: Sample−based surveillance report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    report_context(c,section5_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()

def section6_nohospital(lst_pagenumber=pagenumber_ava_6, lastpage="47", today=date.today().strftime("%d %b %Y")):
    section6_page1_1_1 = "Not applicable because hospital_admission_data.csv file is not available, or in−hospital outcome (in hospital_admission_data.csv file) is not available."
    section6_page1_1 = [section6_page1_1_1]  
    ######### SECTION6: PAGE1 #########
    report_title(c,'Report [6] Mortality in AMR',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'antimicrobial−susceptible infections',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    report_context(c,section6_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()

def annexA_nomicro(lst_pagenumber=pagenumber_ava_annexA, lastpage="47", today=date.today().strftime("%d %b %Y")):
    annexA_page1_1_1 = "Supplementary report on notifiable bacterial disease is not applicable because microbiology_data.xlsx file is not available."
    annexA_page1_1 = [annexA_page1_1_1]
    ######### ANNEXA: PAGE1 #########
    report_title(c,'Annex A: Supplementary report on notifiable bacterial',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_title(c,'diseases',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    report_context(c,annexA_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()

def annexA_nohospital(lst_pagenumber=pagenumber_ava_annexA, lastpage="47", today=date.today().strftime("%d %b %Y")):
    ##Page1
    annexA_page1_1_1 = "Mortality involving the notifiable bacterial diseases is not applicable because hospital_admission_data.csv file is not available, or in−hospital outcome (in hospital_admission_data.csv file) is not available."
    annexA_page1_1 = [annexA_page1_1_1]
    ######### ANNEXA: PAGE3 #########
    report_title(c,'Annex A2: Mortality involving notifiable bacterial infections',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_context(c,annexA_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[2] + " of " + lastpage)
    c.showPage()

def annexB_nomicro(lst_pagenumber=pagenumber_ava_annexB, lastpage="47", today=date.today().strftime("%d %b %Y")):
    annexB_page1_1_1 = "Supplementary report on data indicators is not applicable because microbiology_data.xlsx is not available, list_of_indicators.xlsx is not available, or number of observation is not estimated."
    annexB_page1_1 = [annexB_page1_1_1]
    ######### ANNEXB: PAGE1 #########
    report_title(c,"Annex B: Supplementary report on data indicators",1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    report_context(c,annexB_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    report_todaypage(c,55,30,"Created on: "+today)
    report_todaypage(c,270,30,"Page " + lst_pagenumber[0] + " of " + lastpage)
    c.showPage()

c = canvas.Canvas("./"+"AMR_surveillance_report.pdf")
if check_config(config, "amr_surveillance_section1"):
    try:
        cover(section1_result=sec1_res)
        generatedby(section1_result=sec1_res)
    except Exception as e:
        logger.exception(e)
        pass
tableofcontent()
introduction()
if check_config(config, "amr_surveillance_section1"):
    try:
        section1(sec1_res, sec1_T)
    except Exception as e:
        logger.exception(e)
        pass
if check_config(config, "amr_surveillance_section2"):
    if checkpoint(path_result + sec2_res_i):
        try:
            section2(sec2_res, sec2_merge, 
                    sec2_org1, sec2_org2, sec2_org3, sec2_org4, sec2_org5, sec2_org6, sec2_org7, sec2_org8, 
                    lst_org_format)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        section1to3_nomicro()
if check_config(config, "amr_surveillance_section3"):
    if checkpoint(path_result + sec3_res_i):
        try:
            section3(sec3_res, sec3_pat_val, 
                    sec3_org1_com, sec3_org1_hos, sec3_org2_com, sec3_org2_hos, sec3_org3_com, sec3_org3_hos, sec3_org4_com, sec3_org4_hos, 
                    sec3_org5_com, sec3_org5_hos, sec3_org6_com, sec3_org6_hos, sec3_org7_com, sec3_org7_hos, sec3_org8_com, sec3_org8_hos, 
                    lst_org_format, sec3_lst_numpat, lst_org_short)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        section3_nohospital()
if check_config(config, "amr_surveillance_section4"):
    if checkpoint(path_result + sec4_res_i):
        try:
            section4(sec4_res, sec4_blo_1, sec4_pat_1)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        section4_nohospital()
if check_config(config, "amr_surveillance_section5"):
    if checkpoint(path_result + sec5_res_i):
        try:
            section5(sec5_res, sec5_com_1, sec5_hos_1, sec5_com_amr_1, sec5_hos_amr_1)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        section5_nohospital()
if check_config(config, "amr_surveillance_section6"):
    if checkpoint(path_result + sec6_res_i):
        try:
            section6(sec6_res, sec6_mor_all, 
                    sec6_mor_com_1, sec6_mor_hos_1, sec6_mor_com_2, sec6_mor_hos_2, sec6_mor_com_3, sec6_mor_hos_3, sec6_mor_com_4, sec6_mor_hos_4, 
                    sec6_mor_com_5, sec6_mor_hos_5, sec6_mor_com_6, sec6_mor_hos_6, sec6_mor_com_7, sec6_mor_hos_7, sec6_mor_com_8, sec6_mor_hos_8,
                    lst_org_format, lst_org_short, lst_org_full, sec6_numpat_com, sec6_numpat_hos)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        section6_nohospital()
if check_config(config, "amr_surveillance_annexA"):
    if checkpoint(path_result + secA_res_i):
        try:
            annexA_page1to2(secA_res, annexA_org_page1, secA_pat2)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        annexA_nomicro()
    if checkpoint(path_result + secA_mor_i):
        try:
            annexA_page3(secA_mortal3)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        annexA_nohospital()
if check_config(config, "amr_surveillance_annexB"):
    if checkpoint(path_result + secB_blo_i):
        try:
            annexB(secB_blo_1, secB_blo_bymonth)
        except Exception as e:
            logger.exception(e)
            pass
    else:
        annexB_nomicro()
method()
investor()
c.save()