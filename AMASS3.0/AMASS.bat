@ECHO OFF

del ".\Report_with_patient_identifiers\*.xlsx"
del ".\Report_with_patient_identifiers\*.pdf"
del ".\ResultData\*.xlsx"
del ".\ResultData\*.csv"
del ".\Variables\*.csv"

del ".\AMR_surveillance_report.pdf"
del ".\microbiology_data_reformatted.xlsx"
del ".\data_verification_logfile_report.pdf"
del ".\error_log.txt"
del ".\log_amr_analysis.txt"

rmdir ".\Report_with_patient_identifiers\"
rmdir ".\ResultData\"
rmdir ".\Variables\"
mkdir Report_with_patient_identifiers
mkdir ResultData
mkdir Variables

echo Start Preprocessing: %date% %time%
.\Programs\Python-Portable\Portable_Python-3.7.9\App\Python\python.exe -W ignore .\Programs\AMASS_preprocess\AMASS_preprocess_version_2.py
.\Programs\Python-Portable\Portable_Python-3.7.9\App\Python\python.exe -W ignore .\Programs\AMASS_preprocess\AMASS_preprocess_whonet_version_2.py
echo Start AMR analysis
.\Programs\Python-Portable\Portable_Python-3.7.9\App\Python\python.exe -W ignore .\Programs\AMASS_amr\AMASS3.0_amr_analysis.py
echo Start generating "Data verificator logfile report"
.\Programs\Python-Portable\Portable_Python-3.7.9\App\Python\python.exe -W ignore .\Programs\AMASS_logfile\AMASS_logfile_version_2.py
.\Programs\Python-Portable\Portable_Python-3.7.9\App\Python\python.exe -W ignore .\Programs\AMASS_logfile\AMASS_logfile_err_version_2.py
del ".\error_analysis*.txt"
del ".\error_report*.txt"
del ".\error_logfile_amass.txt"
del ".\error_preprocess.txt"
del ".\error_preprocess_whonet.txt"
del ".\Report_with_patient_identifiers\Report_with_patient_identifiers_annexA_withstatus.xlsx"
del ".\Report_with_patient_identifiers\Report_with_patient_identifiers_annexB_withstatus.xlsx"
echo Finish running AMASS: %date% %time%