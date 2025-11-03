from ZohoCRMAutomatedAuth import ZohoCRMAutomatedAuth
from helper import excel_to_json, assign_sales_person_to_areas, separate_and_store_temp, assgin_leads_to_lead_name

def lead_import(file_path):
    try: 
        crm = ZohoCRMAutomatedAuth()    
        if crm.test_api_connection():
            crm.get_module_fields()      
            file_path = separate_and_store_temp(file_path, True)            
            temp_path = assign_sales_person_to_areas(
                excel_file_path=file_path,
                area_column_name="Area Name", 
                sales_person_column_name="Sales Person"
            )            
            records = excel_to_json(temp_path)
            if records: 
                cmda_success = crm.push_records_to_zoho(records)
                leads_success = assgin_leads_to_lead_name(temp_path, crm)                
                if cmda_success and leads_success:
                    return {
                        "message": "Records pushed to CMDA and Leads created successfully!",
                        "statusCode": 200,
                        "status": True,
                    }
                else:
                    error_msg = []
                    if not cmda_success:
                        error_msg.append("Failed to push some CMDA records")
                    if not leads_success:
                        error_msg.append("Failed to create some Leads")
                    
                    return {
                        "message": "; ".join(error_msg),
                        "statusCode": 400,  
                        "status": False,
                    }
            else:
                return {
                    "message": "No records found in Excel file",
                    "statusCode": 400,  
                    "status": False,
                }
        else:
            return {
                "message": "API connection failed!",
                "statusCode": 400,  
                "status": False,
                "data": [{}]
            }
    except Exception as e:
        return {
            "message": str(e),
            "statusCode": 400,
            "status": False,
            "data": [{}]
        }



