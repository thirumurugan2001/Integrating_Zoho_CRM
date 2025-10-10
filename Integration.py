from ZohoCRMAutomatedAuth import ZohoCRMAutomatedAuth
from helper import excel_to_json,assign_sales_person_to_areas,separate_and_store_temp

def lead_import(file_path):
    try : 
        crm = ZohoCRMAutomatedAuth()    
        if crm.test_api_connection():
            crm.get_module_fields()
            file_path = separate_and_store_temp(file_path)
            temp_path = assign_sales_person_to_areas(excel_file_path=file_path,area_column_name="Area Name", sales_person_column_name="Sales Person")
            records = excel_to_json(temp_path)   
            if records:
                success = crm.push_records_to_zoho(records)
                if success:
                    return {
                        "message": "Records pushed successfully!",
                        "statusCode": 200,
                        "status": True,
                    }
                else:
                    return {
                        "message": "Failed to push some or all records",
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