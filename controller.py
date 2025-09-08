from Integration import lead_import

def lead_validation(data):
    try:
        if isinstance(data, list): 
            if len(data) > 0:
                return lead_import(data)
            else:
                return {
                    "message": "The list you provided is empty. Please ensure it is not empty.",
                    "statusCode": 400,  
                    "status": False,
                    "data": [{}]
                }
        else:
            return {
                "message": "Leads should be in list format.",
                "statusCode": 400,   # ✅ corrected typo
                "status": False,
            }

    except Exception as e:
        return {
            "message": str(e),
            "statusCode": 400,
            "status": False,
        }
