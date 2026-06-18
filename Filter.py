import pandas as pd

#filter the compaines for further investigation(no apps or historical apps) or currnet having app(s) 
input_file_path = "/Users/tianyuzhou/Documents/Finance_RA/pitchbook_wrds.csv"
output_file_path = "/Users/tianyuzhou/Documents/Finance_RA/pitchbook_app_pilot.xlsx"
df = pd.read_csv(input_file_path, usecols= ["CompanyID", "CompanyName", "Description","Website","PrimaryIndustryGroup","PrimaryIndustryCode"])

pilot = df.head(1000).copy()
pilot["Current_App_Status"] = ""
pilot["iOS_App_URL"] = ""
pilot["Android_app_URL"]=""
pilot["Developer_page_URL"] = ""
pilot["Developer_Name"] = ""
pilot["Evidence_note"]=""
pilot["Needs_Historical_review"]=""

pilot.to_excel(output_file_path, index= False)

print("Saved to:",output_file_path)
