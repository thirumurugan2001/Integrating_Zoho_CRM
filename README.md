# 🚀 Zoho CRM Lead Integration with FastAPI  

This project provides a **FastAPI-based microservice** to create leads in **Zoho CRM** automatically.  
It uses **Zoho OAuth2 Automated Authentication with Selenium** and **REST APIs** for secure lead import.  

---

## 📂 Project Structure  

```
.
├── main.py                 # FastAPI application entry point
├── model.py                # Pydantic models (request validation)
├── controller.py           # Lead validation and routing logic
├── integration.py          # Calls Zoho CRM lead importer
├── ZohoCRMAutomatedAuth.py # Zoho authentication & lead creation logic
├── requirements.txt        # Project dependencies
└── README.md               # Documentation
```

---

## ⚙️ Features  

- ✅ **FastAPI backend** with REST API endpoint  
- ✅ **Zoho OAuth2 automated login flow** (with Selenium)  
- ✅ **Token refresh & storage mechanism**  
- ✅ **Bulk lead creation in Zoho CRM**  
- ✅ **Validation for input payloads**  
- ✅ **Docker-friendly & production ready**  

---

## 📦 Installation  

### 1️⃣ Clone the Repository  
```bash
git clone https://github.com/thirumurugan2001/Integrating_Zoho_CRM.git
cd Integrating_Zoho_CRM
```

### 2️⃣ Create Virtual Environment & Install Dependencies  
```bash
python -m venv venv
source venv/bin/activate   # On Linux/Mac
venv\Scripts\activate      # On Windows

pip install -r requirements.txt
```

---

## 🔑 Environment Variables  

Create a **.env** file in the root directory with the following values:  

```ini
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
REDIRECT_URL=https://www.google.com   # or your redirect URI
ORG_ID=your_org_id
EMAIL_ADDRESS=your_email@domain.com
PASSWORD=your_password
AUTH_URL=https://accounts.zoho.com/oauth/v2/auth
TOKEN_URL=https://accounts.zoho.com/oauth/v2/token
API_BASE_URL=https://www.zohoapis.com/crm/v2
TOKEN_FILE_NAME=tokens.json
```

---

## ▶️ Running the Application  

### Start FastAPI Server
```bash
python main.py
```

By default, the app runs at:  
👉 **http://0.0.0.0:8080**

---

## 📡 API Endpoints  

### **1. Create Leads**  
**POST** `/api/create_leads`  

#### ✅ Request Body Example
```json
{
  "leads": [
    {
      "Lead_Name": "Thangamani",
      "Company": "Test Company 1",
      "Email": "john.doe@testcompany1.com",
      "Phone": "+1234567890",
      "Lead_Source": "Website"
    },
    {
      "Lead_Name": "Thangamani",
      "Company": "Test Company 2",
      "Email": "jane.smith@testcompany2.com",
      "Phone": "+1234567891",
      "Lead_Source": "Referral"
    }
  ]
}
```

#### 🔄 Sample Response
```json
{
  "message": "Successfully Created leads",
  "statusCode": 200,
  "status": true,
  "data": [
    {
      "code": "SUCCESS",
      "details": {
        "id": "1234567890001"
      },
      "message": "record added",
      "status": "success"
    }
  ]
}
```

> ⚠️ Note: Currently your `integration.py` returns `statusCode: 400` even for success.  
You may want to update it to return `200` and `status: true` on success.  

---

## 🛠️ Development Notes  

- **Authentication Flow**:  
  - First run triggers **Zoho OAuth2 automated login (via Selenium)**.  
  - Access & Refresh tokens are stored in a local JSON file (`tokens.json`).  
  - Tokens are auto-refreshed when expired.  

- **Lead Validation**:  
  - Ensures leads are passed as a **non-empty list**.  
  - Returns proper error messages if the payload is invalid.  

- **Batch Processing**:  
  - Leads are created in batches of **100 records** for Zoho CRM API compliance.  

---

## 🐳 Docker Support (Optional)  

Create a `Dockerfile`:  

```dockerfile
FROM python:3.10-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["python", "main.py"]
```

Build & Run:  
```bash
docker build -t zoho-fastapi .
docker run -p 8080:8080 --env-file .env zoho-fastapi
```

---

## 🚨 Error Handling  

- **Invalid List Input**  
```json
{
  "message": "Leads should be in list format.",
  "statusCode": 400,
  "status": false
}
```

- **Empty List Input**  
```json
{
  "message": "The list you provided is empty. Please ensure it is not empty.",
  "statusCode": 400,
  "status": false,
  "data": [{}]
}
```

- **Zoho API Failure**  
```json
{
  "message": "Zoho CRM API failed to create leads",
  "statusCode": 500,
  "status": false
}
```

---

## 📜 License  

MIT License – feel free to use and modify.  
