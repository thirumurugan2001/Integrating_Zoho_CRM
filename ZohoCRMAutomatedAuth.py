# ZohoCRMAutomatedAuth.py
import requests
import json
import time
import os
from datetime import datetime, timedelta
from urllib.parse import urlparse, parse_qs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException, ElementNotInteractableException
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from dotenv import load_dotenv
load_dotenv()


class ZohoCRMAutomatedAuth: 

    def __init__(self):
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.redirect_uri = os.getenv("REDIRECT_URL")
        self.org_id = os.getenv("ORG_ID")
        self.email = os.getenv("EMAIL_ADDRESS")
        self.password = os.getenv("PASSWORD")
        self.auth_url = os.getenv("AUTH_URL")
        self.token_url = os.getenv("TOKEN_URL")
        self.api_base_url = os.getenv("API_BASE_URL")
        self.zoho_model_name = os.getenv("ZOHO_MODEL_NAME")
        self.access_token = None
        self.refresh_token = None
        self.token_expires_at = None
        self.token_file = os.getenv("TOKEN_FILE_NAME")
    
    def setup_driver(self, headless=False):
        chrome_options = Options()
        if headless:
            chrome_options.add_argument("--headless=new")        
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        chrome_options.add_argument("--disable-images")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--start-maximized")        
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")        
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            driver.implicitly_wait(10)
            return driver
        except WebDriverException as e:
            print(f"Error setting up driver: {e}")
            return None
        
    def wait_and_find_element(self, driver, selectors, timeout=30):
        wait = WebDriverWait(driver, timeout)
        for selector_type, selector_value in selectors:
            try:
                if selector_type == "id":
                    element = wait.until(EC.element_to_be_clickable((By.ID, selector_value)))
                elif selector_type == "name":
                    element = wait.until(EC.element_to_be_clickable((By.NAME, selector_value)))
                elif selector_type == "xpath":
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, selector_value)))
                elif selector_type == "css":
                    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector_value)))
                if element.is_displayed() and element.is_enabled():
                    return element, selector_type, selector_value
            except (TimeoutException, ElementNotInteractableException):
                continue
        return None, None, None
    
    def safe_click(self, driver, element, description="element"):
        try:
            element.click()
            return True
        except ElementNotInteractableException:
            try:
                driver.execute_script("arguments[0].click();", element)
                return True
            except Exception as e:
                try:
                    ActionChains(driver).move_to_element(element).click().perform()
                    return True
                except Exception as e:
                    print(f"Failed to click {description}: {e}")
                    return False
    
    def safe_send_keys(self, driver, element, text, description="field"):
        try:
            element.clear()
            time.sleep(0.5)
            element.send_keys(text)
            return True
        except Exception as e:
            try:
                driver.execute_script("arguments[0].value = '';", element)
                driver.execute_script("arguments[0].value = arguments[1];", element, text)
                return True
            except Exception as e:
                print(f"Failed to enter text in {description}: {e}")
                return False
    
    def handle_tfa_banner_page(self, driver):
        current_url = driver.current_url
        if "tfa-banner" in current_url or "announcement" in current_url:
            continue_selectors = [
                ("xpath", "//button[contains(text(), 'Continue')]"),
                ("xpath", "//button[contains(text(), 'Skip')]"),
                ("xpath", "//button[contains(text(), 'Later')]"),
                ("xpath", "//button[contains(text(), 'Not now')]"),
                ("xpath", "//a[contains(text(), 'Continue')]"),
                ("xpath", "//a[contains(text(), 'Skip')]"),
                ("xpath", "//a[contains(text(), 'Later')]"),
                ("xpath", "//input[@value='Continue']"),
                ("xpath", "//input[@value='Skip']"),
                ("id", "continue"),
                ("id", "skip"),
                ("id", "later"), 
                ("css", ".continue-btn"),
                ("css", ".skip-btn"),
                ("css", "button.primary"),
                ("css", "a.primary"),
                ("xpath", "//button[contains(@class, 'primary')]"),
                ("xpath", "//a[contains(@class, 'primary')]")
            ]
            element, _, _ = self.wait_and_find_element(driver, continue_selectors, 20)
            if element:
                if self.safe_click(driver, element, "continue/skip button"):
                    time.sleep(3)
                    return True
            else:
                parsed_url = urlparse(current_url)
                query_params = parse_qs(parsed_url.query)
                service_url = query_params.get('serviceurl', [None])[0]
                if service_url:
                    from urllib.parse import unquote
                    decoded_service_url = unquote(service_url)
                    driver.get(decoded_service_url)
                    time.sleep(3)
                    return True
        return False
    
    def get_authorization_url(self):
        params = {
            'scope': 'ZohoCRM.modules.ALL,ZohoCRM.settings.ALL',
            'client_id': self.client_id,
            'response_type': 'code',
            'access_type': 'offline',
            'redirect_uri': self.redirect_uri
        }
        auth_url = f"{self.auth_url}?" + "&".join([f"{k}={v}" for k, v in params.items()])
        return auth_url
    
    def automate_oauth_flow(self, headless=False):
        driver = self.setup_driver(headless)
        if not driver:
            return False
        try:
            auth_url = self.get_authorization_url()
            driver.get(auth_url)
            time.sleep(3)
            email_selectors = [
                ("id", "login_id"),
                ("name", "LOGIN_ID"),
                ("xpath", "//input[@type='email']"),
                ("xpath", "//input[@placeholder='Email ID']"),
                ("css", "input[type='email']"),
                ("xpath", "//input[contains(@class, 'email')]")
            ]
            email_element, _, _ = self.wait_and_find_element(driver, email_selectors, 30)
            if not email_element:
                self.debug_page(driver)
                return False            
            if not self.safe_send_keys(driver, email_element, self.email, "email field"):
                return False
            next_selectors = [
                ("id", "nextbtn"),
                ("id", "signin_submit"),
                ("xpath", "//button[contains(text(), 'Next')]"),
                ("xpath", "//button[contains(text(), 'Continue')]"),
                ("xpath", "//input[@value='Next']"),
                ("xpath", "//input[@value='Continue']"),
                ("css", "button[type='submit']"),
                ("xpath", "//button[@type='submit']")
            ]
            next_element, _, _ = self.wait_and_find_element(driver, next_selectors, 20)
            if next_element:
                if self.safe_click(driver, next_element, "next button"):
                    time.sleep(3)
            else:
                print("⚠️ Next button not found, continuing...")            
            password_selectors = [
                ("id", "password"),
                ("name", "PASSWORD"),
                ("xpath", "//input[@type='password']"),
                ("css", "input[type='password']"),
                ("xpath", "//input[contains(@class, 'password')]")
            ]
            password_element, _, _ = self.wait_and_find_element(driver, password_selectors, 30)
            if not password_element:
                self.debug_page(driver)
                return False
            if not self.safe_send_keys(driver, password_element, self.password, "password field"):
                return False            
            signin_selectors = [
                ("id", "nextbtn"),
                ("id", "signin_submit"),
                ("xpath", "//button[contains(text(), 'Sign in')]"),
                ("xpath", "//button[contains(text(), 'Sign In')]"),
                ("xpath", "//button[contains(text(), 'Login')]"),
                ("xpath", "//input[@value='Sign in']"),
                ("xpath", "//input[@value='Sign In']"),
                ("css", "button[type='submit']"),
                ("xpath", "//button[@type='submit']")
            ]
            signin_element, _, _ = self.wait_and_find_element(driver, signin_selectors, 20)
            if not signin_element:
                self.debug_page(driver)
                return False
            if not self.safe_click(driver, signin_element, "sign in button"):
                return False
            time.sleep(5)            
            if self.handle_tfa_banner_page(driver):
                time.sleep(3)            
            accept_selectors = [
                ("xpath", "//button[contains(text(), 'Accept')]"),
                ("xpath", "//button[contains(text(), 'Allow')]"),
                ("xpath", "//button[contains(text(), 'Authorize')]"),
                ("xpath", "//input[@value='Accept']"),
                ("xpath", "//input[@value='Allow']"),
                ("id", "accept"),
                ("id", "allow"),
                ("css", "button.accept"),
                ("css", "button.allow"),
                ("xpath", "//button[contains(@class, 'accept')]"),
                ("xpath", "//button[contains(@class, 'allow')]")
            ]
            accept_element, _, _ = self.wait_and_find_element(driver, accept_selectors, 15)
            if accept_element:
                if self.safe_click(driver, accept_element, "accept button"):
                    time.sleep(3)
            else:
                print("No authorization page found or already authorized")            
            max_wait_time = 60
            start_time = time.time()
            while time.time() - start_time < max_wait_time:
                current_url = driver.current_url
                if "google.com" in current_url and "code=" in current_url:
                    break
                if "tfa-banner" in current_url or "announcement" in current_url:
                    self.handle_tfa_banner_page(driver)
                time.sleep(2)
            current_url = driver.current_url
            if "google.com" in current_url and "code=" in current_url:
                parsed_url = urlparse(current_url)
                code = parse_qs(parsed_url.query).get('code', [None])[0]                
                if code:
                    print("Authorization code obtained, getting access token...")
                    success = self.get_access_token(code)
                    return success
                else:
                    return False
            else:
                self.debug_page(driver)
                return False
        except Exception as e:
            self.debug_page(driver)
            return False
        finally:
            driver.quit()
    
    def debug_page(self, driver):
        try:
            elements_info = []
            for tag in ['input', 'button', 'a']:
                elements = driver.find_elements(By.TAG_NAME, tag)
                for element in elements[:10]:
                    try:
                        info = {
                            'tag': tag,
                            'id': element.get_attribute('id'),
                            'name': element.get_attribute('name'),
                            'type': element.get_attribute('type'),
                            'class': element.get_attribute('class'),
                            'text': element.text[:50] if element.text else '',
                            'displayed': element.is_displayed(),
                            'enabled': element.is_enabled()
                        }
                        elements_info.append(info)
                    except:
                        continue
            for info in elements_info:
                print(f"  {info}")
        except Exception as e:
            print(f"Debug error: {e}")
    
    def get_access_token(self, authorization_code):        
        data = {
            'grant_type': 'authorization_code',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'redirect_uri': self.redirect_uri,
            'code': authorization_code
        }
        try:
            response = requests.post(self.token_url, data=data)
            if response.status_code == 200:
                token_data = response.json()
                self.access_token = token_data['access_token']
                self.refresh_token = token_data['refresh_token']
                expires_in = token_data.get('expires_in', 3600)
                self.token_expires_at = datetime.now() + timedelta(seconds=expires_in)
                self.save_tokens()
                return True
            else:
                print(f"Failed to get access token: {response.status_code} - {response.text}")
                return False
        except Exception as e:
            print(f"Error getting access token: {e}")
            return False
    
    def refresh_access_token(self):
        if not self.refresh_token:
            return False
        data = {
            'grant_type': 'refresh_token',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'refresh_token': self.refresh_token
        }
        try:
            response = requests.post(self.token_url, data=data)
            if response.status_code == 200:
                token_data = response.json()
                self.access_token = token_data['access_token']
                expires_in = token_data.get('expires_in', 3600)
                self.token_expires_at = datetime.now() + timedelta(seconds=expires_in)
                self.save_tokens()
                print("Access token refreshed successfully!")
                return True
            else:
                print(f"Failed to refresh token: {response.status_code} - {response.text}")
                return False
        except Exception as e:
            print(f"Error refreshing token: {e}")
            return False
    
    def save_tokens(self):
        token_data = {
            'access_token': self.access_token,
            'refresh_token': self.refresh_token,
            'expires_at': self.token_expires_at.isoformat() if self.token_expires_at else None,
            'client_id': self.client_id
        }
        try:
            with open(self.token_file, 'w') as f:
                json.dump(token_data, f, indent=2)
        except Exception as e:
            print(f"Error saving tokens: {e}")

    def load_tokens(self):
        if os.path.exists(self.token_file):
            try:
                with open(self.token_file, 'r') as f:
                    token_data = json.load(f)
                self.access_token = token_data.get('access_token')
                self.refresh_token = token_data.get('refresh_token')
                expires_at_str = token_data.get('expires_at')
                if expires_at_str:
                    self.token_expires_at = datetime.fromisoformat(expires_at_str)                
                return True
            except Exception as e:
                return False
        return False
    
    def ensure_valid_token(self):
        if not self.access_token:
            self.load_tokens()  
        if self.token_expires_at and datetime.now() >= self.token_expires_at:
            if not self.refresh_access_token():
                return self.automate_oauth_flow()
            return True        
        if not self.access_token:
            return self.automate_oauth_flow()
        return True
    
    def format_record_for_zoho(self, record):
        formatted_record = {}
        field_mapping = {
            "Lead_Owner": "Name",
            "Email ID": "Email",
            "Mobile No.": "Mobile_Number",
            "Date of permit": "Date_of_Permit",
            "Applicant Name": "Applicant_Name",
            "Nature of Development": "Nature_of_Developments",
            "Dwelling Unit Info": "Dwelling_Unit_Info",
            "Sales Person": "Lead_Source",
            "Lead_Name": "Lead_Name",
            "Reference": "Reference",
            "Company_Name": "Company_Name",
            "Architect Name": "Architect",
            "Planning Permission No.": "Plan_Permission",
            "Applicant Address": "Applicant_Address",
            "Future_Projects": "Future_Project", 
            "Creation_Time": "Creation_Time",
            "Which_Brand_Looking_for": "Which_Brand_Looking_for",
            "How_Much_Square_Feet": "How_Much_Square_Feet",
            "Area Name": "Area_Name",  
            "Site Address": "Site_Address"
        }        
        try:
            dwelling_units = record.get("Dwelling Unit Info", 0)
            if dwelling_units and not pd.isna(dwelling_units):
                bathrooms = int(float(dwelling_units)) * 2
                formatted_record["No_of_bathrooms"] = str(bathrooms)
            else:
                formatted_record["No_of_bathrooms"] = "0"
        except (ValueError, TypeError):
            formatted_record["No_of_bathrooms"] = "0"
        
        for excel_field, zoho_field in field_mapping.items():
            if excel_field not in record or record[excel_field] is None or pd.isna(record[excel_field]):
                formatted_record[zoho_field] = ""
                continue
            
            value = record[excel_field]
            
            if excel_field in ["Creation_Time", "Date_of_Permit"]:
                if isinstance(value, str):
                    try:
                        dt = datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
                        formatted_record[zoho_field] = dt.strftime("%Y-%m-%dT%H:%M:%S+05:30")
                    except ValueError:
                        try:
                            dt = datetime.strptime(value, "%Y-%m-%d")
                            formatted_record[zoho_field] = dt.strftime("%Y-%m-%d")
                        except ValueError:
                            formatted_record[zoho_field] = str(value)
                elif hasattr(value, "strftime"):
                    formatted_record[zoho_field] = value.strftime("%Y-%m-%dT%H:%M:%S+05:30")
                else:
                    formatted_record[zoho_field] = str(value)
            elif excel_field in ["Dwelling Unit Info", "How_Much_Square_Feet"]:
                try:
                    formatted_record[zoho_field] = str(int(float(value)))
                except (ValueError, TypeError):
                    formatted_record[zoho_field] = "0"
            elif excel_field == "Email ID":
                email_str = str(value).strip()
                if "@" in email_str and "." in email_str:
                    formatted_record[zoho_field] = email_str
                else:
                    formatted_record[zoho_field] = "" 
            elif excel_field == "Mobile No.":
                mobile_str = str(int(float(value))) if isinstance(value, (int, float)) else str(value).strip()
                formatted_record[zoho_field] = mobile_str
            else:
                formatted_record[zoho_field] = str(value).strip()
        
        # Set the Name field
        if record.get("Applicant Name") and not pd.isna(record.get("Applicant Name")):
            formatted_record["Name"] = str(record["Applicant Name"]).strip()
        elif record.get("Lead_Name") and not pd.isna(record.get("Lead_Name")):
            formatted_record["Name"] = str(record["Lead_Name"]).strip()
        elif record.get("Company_Name") and not pd.isna(record.get("Company_Name")):
            formatted_record["Name"] = str(record["Company_Name"]).strip()
        else:
            formatted_record["Name"] = f"Record_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        
        return formatted_record

    def push_records_to_zoho(self, records, batch_size=100):
        """Push records to Zoho CRM in batches"""
        if not self.ensure_valid_token():
            print("Failed to ensure valid token")
            return False
        if not records:
            print("No records to push")
            return True 
        total_records = len(records)
        successful_records = 0
        failed_records = 0
        
        for i in range(0, total_records, batch_size):
            batch = records[i:i + batch_size]
            formatted_batch = []
            for record in batch:
                formatted_record = self.format_record_for_zoho(record)
                if formatted_record:
                    formatted_batch.append(formatted_record)
            if not formatted_batch:
                continue

            print(json.dumps({'data': formatted_batch}, indent=2))
            
            url = f"{self.api_base_url}/{self.zoho_model_name}"
            headers = {
                'Authorization': f'Zoho-oauthtoken {self.access_token}',
                'Content-Type': 'application/json'
            }
            payload = {
                'data': formatted_batch,
                'trigger': ['approval', 'workflow', 'blueprint']
            }
            try:
                response = requests.post(url, json=payload, headers=headers)
                if response.status_code == 201:
                    response_data = response.json()
                    batch_success = 0
                    batch_failed = 0
                    if 'data' in response_data:
                        for result in response_data['data']:
                            if result.get('status') == 'success':
                                batch_success += 1
                            else:
                                batch_failed += 1
                                print(f"❌ Record failed: {result.get('message', 'Unknown error')}")
                                print(f"   Details: {result.get('details', 'No details')}")
                    successful_records += batch_success
                    failed_records += batch_failed
                else:
                    print(f"❌ HTTP Error {response.status_code}: {response.text}")
                    failed_records += len(formatted_batch)
                if i + batch_size < total_records:
                    time.sleep(1)  
            except Exception as e:
                print(f"❌ Exception while pushing batch: {e}")
                import traceback
                traceback.print_exc()
                failed_records += len(formatted_batch)
        
        print(f"\n✅ Push completed: {successful_records} successful, {failed_records} failed out of {total_records} total")
        return successful_records > 0
      
    def test_api_connection(self):
        if not self.ensure_valid_token():
            print("Failed to ensure valid token")
            return False
        url = f"{self.api_base_url}/settings/modules"
        headers = {
            'Authorization': f'Zoho-oauthtoken {self.access_token}',
            'Content-Type': 'application/json'
        }
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                modules = response.json()
                module_names = [module['api_name'] for module in modules.get('modules', [])]
                print(f"Available modules: {module_names}")
                if self.zoho_model_name in module_names:
                    print(f"✅ Model '{self.zoho_model_name}' found!")
                    return True
                else:
                    print(f"❌ Model '{self.zoho_model_name}' not found in available modules")
                    return False
            else:
                print(f"Failed to get modules: {response.status_code} - {response.text}")
                return False
        except Exception as e:
            print(f"Error testing API connection: {e}")
            return False

    def get_module_fields(self):
        if not self.ensure_valid_token():
            print("Failed to ensure valid token")
            return False
        url = f"{self.api_base_url}/settings/fields?module={self.zoho_model_name}"
        headers = {
            'Authorization': f'Zoho-oauthtoken {self.access_token}',
            'Content-Type': 'application/json'
        }
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                fields_data = response.json()
                fields = fields_data.get('fields', [])                
                for field in fields:
                    field_name = field.get('api_name', 'Unknown')
                    field_label = field.get('field_label', 'No Label')
                    field_type = field.get('data_type', 'Unknown')
                    required = field.get('required', False)                    
                    status = "✅" if required else "⚪"
                return True
            else:
                return False
        except Exception as e:
            print(f"Error getting module fields: {e}")
            return False
        
# helper.py
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
            crm_columns = [
                "Email ID", "Mobile No.", "Date of permit", "Site Address",
                "Applicant Name", "Nature of Development", "Dwelling Unit Info","Area Name", 
                "Sales Person", "Lead_Name", "Reference", "No_of_bathrooms", 
                "Company_Name", "Architect Name", "Planning Permission No.", 
                "Applicant Address", "Future_Projects", "Creation_Time", 
                "Which_Brand_Looking_for", "How_Much_Square_Feet"
            ]
            print("\n🔍 CRM Column Check:")
            for crm_col in crm_columns:
                if crm_col in sample_record:
                    print(f"✅ {crm_col}: Found in Excel")
                else:
                    print(f"❌ {crm_col}: Missing from Excel")
        return cleaned_records        

    except Exception as e:
        print(f"Error in excel_to_json: {str(e)}")
        return []