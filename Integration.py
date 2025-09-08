from ZohoCRMAutomatedAuth import ZohoCRMLeadImporter

def lead_import(sample_leads):
    try :   
        importer = ZohoCRMLeadImporter()
        importer.import_leads_from_list(sample_leads)
        return {
            "message": "Successfully Created leads",
            "statusCode": 400,
            "status": False,
        }
    except Exception as e:
        return {
            "message": str(e),
            "statusCode": 400,
            "status": False,
        }
