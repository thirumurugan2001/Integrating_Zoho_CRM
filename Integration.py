from ZohoCRMAutomatedAuth import ZohoCRMAutomatedAuth
from helper import excel_to_json

def lead_import(file_path):
    try : 
        crm = ZohoCRMAutomatedAuth()    
        if crm.test_api_connection():
            crm.get_module_fields()        
            records = excel_to_json(file_path)       
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