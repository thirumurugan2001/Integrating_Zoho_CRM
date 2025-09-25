import pandas as pd

def excel_to_json(file_path: str):
    try:
        df = pd.read_excel(file_path)
        records = df.to_dict(orient="records")
        cleaned_records = []
        for record in records:
            cleaned_record = {} 
            for key, value in record.items():
                if pd.notna(value):
                    cleaned_record[key] = value
            cleaned_records.append(cleaned_record)       
        if cleaned_records:
            sample_record = cleaned_records[0]
            crm_columns = [ "Email", "Mobile_Number", "Date_of_Permit", 
                "Applicant_Name", "Nature_of_Development", "Dueling_Units", 
                "Lead_Source", "Lead_Name", "Reference", "No_of_bathrooms", 
                "Company_Name", "Architect", "Plan_Permission", 
                "Applicant_Address", "Future_Projects", "Creation_Time", 
                "Which_Brand_Looking_for", "How_Much_Square_Feet"
            ]
            for crm_col in crm_columns:
                if crm_col in sample_record:
                    print(f"✅ {crm_col}: Found in Excel")
                else:
                    print(f"❌ {crm_col}: Missing from Excel")
        return cleaned_records        
    except Exception as e:
        print(f"Error in excel_to_json: {str(e)}")
        return []