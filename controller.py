from Integration import lead_import
import os

def lead_validation(file_path):
    try:
        if not os.path.exists(file_path):
            return {
                "message": "The file path you provided does not exist. Please check the path and try again.",
                "statusCode": 400,  
                "status": False,
                "data": [{}]
            }
        else:
            return lead_import(file_path)
    except Exception as e:
        return {
            "message": str(e),
            "statusCode": 400,
            "status": False,
        }
    
    
