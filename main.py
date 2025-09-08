from fastapi import FastAPI
import uvicorn
from model import *
from controller import lead_validation
app = FastAPI()

@app.post("/api/create_leads")
async def create_leads(leads: Leads):
    try:
        response = lead_validation(leads.leads)
        return response
    except Exception as e:
        print("Error in create_leads:", str(e))
        return {
            "message": str(e),
            "statusCode": 400,
            "status": False,
            "data": [{}]
        }

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8080, reload=True)
