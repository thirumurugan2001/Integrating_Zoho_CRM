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
import re
import traceback
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
            "Sales Person": "Lead_Owner",  # ✅ Changed: Sales Person maps to Lead_Owner
            "Email ID": "Email",
            "Mobile No.": "Mobile_Number",
            "Date of permit": "Date_of_Permit",
            "Applicant Name": "Lead_Name",  # ✅ Changed: Applicant Name maps to Lead_Name
            "Nature of Development": "Nature_of_Developments",
            "Dwelling Unit Info": "Dwelling_Unit_Info",
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
            dwelling_units = record.get("Dwelling Unit Info")
            if dwelling_units is not None and not pd.isna(dwelling_units):
                dwelling_str = str(dwelling_units).strip()
                if dwelling_str and dwelling_str != '' and dwelling_str.lower() != 'nan':
                    try:
                        numbers = re.findall(r'\d+', dwelling_str)
                        if numbers:
                            dwelling_value = int(numbers[0])
                            bathrooms = dwelling_value * 2
                            formatted_record["No_of_bathrooms"] = str(bathrooms)
                        else:
                            formatted_record["No_of_bathrooms"] = "0"
                    except (ValueError, TypeError) as e:
                        formatted_record["No_of_bathrooms"] = "0"
                else:
                    formatted_record["No_of_bathrooms"] = "0"
            else:
                formatted_record["No_of_bathrooms"] = "0"
        except Exception as e:
            print(f"❌ Error calculating bathrooms: {e}")
            traceback.print_exc()
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
                    numbers = re.findall(r'\d+', str(value))
                    if numbers:
                        formatted_record[zoho_field] = numbers[0]
                    else:
                        formatted_record[zoho_field] = "0"
                except (ValueError, TypeError):
                    formatted_record[zoho_field] = "0"            
            elif excel_field == "Email ID":
                email_str = str(value).strip()
                if "@" in email_str and "." in email_str:
                    formatted_record[zoho_field] = email_str
                else:
                    formatted_record[zoho_field] = "" 
            elif excel_field == "Mobile No.":
                try:
                    mobile_str = str(value).strip()
                    mobile_clean = re.sub(r'[^\d+]', '', mobile_str)
                    formatted_record[zoho_field] = mobile_clean
                except:
                    formatted_record[zoho_field] = str(value).strip()
            else:
                formatted_record[zoho_field] = str(value).strip()
        
        # ✅ Updated: Set "Name" field using Applicant Name as priority
        if record.get("Applicant Name") and not pd.isna(record.get("Applicant Name")):
            formatted_record["Name"] = str(record["Applicant Name"]).strip()
        elif record.get("Company_Name") and not pd.isna(record.get("Company_Name")):
            formatted_record["Name"] = str(record["Company_Name"]).strip()
        else:
            formatted_record["Name"] = f"Record_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        
        if "No_of_bathrooms" in formatted_record:
            pass
        else:
            print(f"⚠️ WARNING: No_of_bathrooms missing from formatted record!")
        
        formatted_record['Lead_Source'] = "Website"
        return formatted_record
    def push_records_to_zoho(self, records, batch_size=100):
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
# Helper.py
import os
import re
import tempfile
import smtplib
from datetime import datetime
from typing import Optional
import pandas as pd
from fuzzywuzzy import fuzz
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

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
                "Sales Person", "Lead_Name", "Reference", "Company_Name", "Architect Name", "Planning Permission No.", 
                "Applicant Address", "Future_Projects", "Creation_Time", 
                "Which_Brand_Looking_for", "How_Much_Square_Feet"
            ]
        return cleaned_records        

    except Exception as e:
        print(f"Error in excel_to_json: {str(e)}")
        return []

def send_unmatched_areas_alert(unmatched_df: pd.DataFrame, original_file_name: str = "input_file.xlsx") -> bool:
    try:
        sender_mailId = os.getenv("SENDER_MAIL", "riverpearlsolutions@gmail.com")
        passKey = os.getenv("APP_PASSWORD", "gwvcgbvjvttpvlja")
        recipient_email = os.getenv("RECIPIENT_MAIL")        
        if not sender_mailId or not passKey:
            print("Error: Email credentials not found")
            return False        
        if unmatched_df.empty:
            print("No unmatched areas to report")
            return True        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", mode='wb')
        unmatched_df.to_excel(temp_file.name, index=False)
        temp_file_path = temp_file.name
        temp_file.close()        
        attachment_filename = f"Unmatched_Areas_{timestamp}.xlsx"        
        msg = MIMEMultipart()
        msg['From'] = sender_mailId
        msg['To'] = recipient_email
        msg['Subject'] = f"Alert: Unmatched Areas Found"     
        total_unmatched = len(unmatched_df)
        unique_areas = unmatched_df['Area Name'].nunique() if 'Area Name' in unmatched_df.columns else 0
        body = f'''
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
                <div style="background-color: #ff6b6b; color: white; padding: 15px; border-radius: 5px;">
                    <h2 style="margin: 0;">⚠️ Unmatched Areas Alert</h2>
                </div>
                
                <div style="padding: 20px; background-color: #f9f9f9; margin-top: 20px; border-radius: 5px;">
                    <p>Dear Team,</p>
                    <p>The system has identified areas that could not be matched to any salesperson during the processing of <strong>{original_file_name}</strong>.</p>
                    
                    <div style="background-color: white; padding: 15px; border-left: 4px solid #ff6b6b; margin: 20px 0;">
                        <h3 style="margin-top: 0; color: #ff6b6b;">Summary</h3>
                        <ul style="list-style: none; padding: 0;">
                            <li>📊 <strong>Total Unmatched Records:</strong> {total_unmatched}</li>
                            <li>📍 <strong>Unique Unmatched Areas:</strong> {unique_areas}</li>
                            <li>📅 <strong>Generated On:</strong> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</li>
                        </ul>
                    </div>
                    
                    <p><strong>Action Required:</strong></p>
                    <ol>
                        <li>Review the attached Excel file containing all unmatched records</li>
                        <li>Update the SALES_PERSON_AREAS mapping in the system if needed</li>
                        <li>Manually assign salespeople to these areas</li>
                        <li>Reprocess the file after updates</li>
                    </ol>
                    
                    <p>The unmatched areas data is attached to this email for your review and action.</p>
                </div>
                
                <hr style="border: none; border-top: 1px solid #ddd; margin: 30px 0;">
                
                <div style="font-size: 12px; color: #666;">
                    <p><strong>VPEARL SOLUTIONS - An AI Company, Chennai</strong></p>
                    <p>
                        🌐 <a href="https://vpearlsolutions.com/" target="_blank" style="color: #4a90e2;">Website</a> | 
                        🔗 <a href="https://www.linkedin.com/company/vpealsoutions/" target="_blank" style="color: #4a90e2;">LinkedIn</a> | 
                        📷 <a href="https://www.instagram.com/vpearl_solutions" target="_blank" style="color: #4a90e2;">Instagram</a> | 
                        📘 <a href="https://www.facebook.com/profile.php?id=61572978223085" target="_blank" style="color: #4a90e2;">Facebook</a>
                    </p>
                    <p style="font-size: 10px; color: #999;">
                        <em>This is an automated alert. Please do not reply to this email.</em>
                    </p>
                </div>
            </div>
        </body>
        </html>
        '''        
        msg.attach(MIMEText(body, 'html'))
        if os.path.exists(temp_file_path):
            with open(temp_file_path, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                attachment.add_header('Content-Disposition', 'attachment', filename=attachment_filename)
                msg.attach(attachment)
        else:
            print(f"Error: Temporary file not found at {temp_file_path}")
            return False
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_mailId, passKey)
            server.sendmail(sender_mailId, recipient_email, msg.as_string())
        try:
            os.unlink(temp_file_path)
        except:
            pass
        return True
    except Exception as e:
        print(f"❌ Error in send_unmatched_areas_alert function: {str(e)}")
        return False

def assign_sales_person_to_areas(excel_file_path: str,area_column_name: str = 'Area Name',sales_person_column_name: str = 'Sales Person',sheet_name: str = None,fuzzy_match_threshold: int = 100) -> dict:
    
    SALES_PERSON_AREAS = {
        "Abhishek": ["Adambakkam","Alandur", "Alandur Guindy", "Guindy", "Madipakkam", 
            "Medavakkam", "Nanganallur", "Pallikaranai", "Thalakananchery", 
            "Thalakkanancheri", "Thalakkananchery", "Thalakkancheri", "Velachery"
        ],
        "Jagan": [
            "Adyar", "Athipattu", "Egmore", "Kottur", "Koyambedu", "koyambedu", 
            "Koyembedu", "Mogappair", "Mullam", "Naduvakarai", "Naduvankarai", 
            "Naduvankkarai", "Nekundram", "Nerkundram", "Nolambur", "Nungambakkam", 
            "Pallipattu", "Part of Thirumangalam", "Periyakudal", "Alwarpet",
            "Secretariat Colony Kilpauk Chennai.", "Urur", "Vada Agaram", "Vepery","Aminjikarai"
        ],
        "Karthik": [
            "Arumbakkam", "Ayyappanthangal", "Ekkaduthangal", "Goparasanallur", 
            "Kalikundram", "Kanagam", "Karambakkam", "Kodambakkam", "Kolapakkam", 
            "Kulamanivakkam", "Madhananthapuram", "Madhandhapuram", "Manapakkam", 
            "Mangadu-B", "Moulivakkam", "Noombal", "Pammal", "Panaveduthottam", 
            "Parivakkam", "Porur", "Puliyur", "Saligramam", "Tharapakkam", 
            "Valasaravakkam", "Virugambakkam", "Voyalanallur-A"
        ],
        "Ventakesh": [
            "Agaramthen", "Anakaputhur", "Chembarambakkam", "Cowl Bazaar", 
            "Gowrivakkam", "Karapakkam", "Kaspapuram", "Kulathuvancheri", 
            "Kundrathur", "Kundrathur - A", "Kundrathur - B", "Kundrathur-A", 
            "Kundrathur-B", "Malayambakkam", "Manancheri", "Mannivakkam", 
            "Meppedu", "Mudichur", "Mullam", "Nandambakkam", "Nanmangalam", 
            "Naduveerapattu", "Nedungundram", "Nedunkundram", "Nemilichery", 
            "Ottiyambakkam", "Palanthandalam", "Pallavaram", "Pallavarm", 
            "Perumbakkam", "Perungalathur", "Rajakilpakkam", "S.Kulathur", 
            "Selaiyur", "Sirukalathur", "Tambaram", "Thirumudivakkam", 
            "Thiruneermalai", "Thiruvancheri", "Vandalur", "Varadarajapuram", 
            "Varadharajapuram", "Vengaivasal", "Vengambakkam", 
            "Ward No.C of Tambaram", "Tambaram"
        ],
        "Dinikaran": [
            "Kottivakkam", "Kovilambakkam", "Neelangarai", "Okkiam Thoraipakkam", 
            "Okkiyam Thoraipakkam", "part of Sholinganallur", "Perungudi", 
            "Sholinganallur", "Thiiruvanmiyur", "Thiruvanmiyur", "Thoraipakkam"
        ],
        "Balachander": [
            "Agraharammel", "Angadu", "Layon Pullion", "Maduravoyal"
        ],
        "Sithalapakkam": [
            "Sithalapakkam"
        ],
        "Jagan / Balachander": [
            "Adayalampattu", "Alamathi", "Ambathur", "Ambattur", "Arumandai", 
            "at Kondakarai Kuruvimedu Panchayat Road and", "at Orakkadu", 
            "at Puzhal", "Ayanambakkam", "Ayanavaram", "Budur", "BUDUR", 
            "Chintadripet", "Girudalapuram", "Kannapalayam", "Karanodai", 
            "Karunakaracheri", "Kathirvedu", "Korattur", "Korattur A", "Kosapur", 
            "Kovilpadagai", "Layon Grant", "Madhavaram", "Mijur", "Minjur", 
            "Minjur II", "Nayar-II", "Nemam", "Oragadam", "Orakkadu", "Padi", 
            "Padiyanallur", "Pakkam", "Palanjur", "Paleripattu", "part of Ayapakkam", 
            "Paruthipattu", "Perambur", "Peravallur", "Periyamullaivoyal", 
            "Perungavur", "Peruvallur", "Ponneri", "Purasaiwalkam", "Purasalwalkam", 
            "Purursawalkkam", "Purusawalkam", "Seemapuram", "Sholavaram", 
            "Sirugavoor", "Sothuperumbedu", "Thirumanam", "Thirunindravur B", 
            "Thiruninravur", "Thiruninravur-A", "Thiruninravur-B", "Thiruvotriyur", 
            "Tondairpet", "Tondiarpet", "Vanagaram", "Vayalanallur", "Vayalanallur-A", 
            "Veeraragavapuram", "Veeraraghavapuram", "Venkatapuram", 
            "Vilangadupakkam", "Villivakkam", "Paruthipattu"
        ],
        "Karthik / Ventakesh": [
            "Gerugambakkam", "Kollacheri", "Kulappakkam", "Kuthambakkam", 
            "Poonamallee", "Rendamkattalai", "Rendankattalai", "Sikkarayapuram", 
            "Vellavedu", "Zamin Pallavaram", "Zamin Pallvaram", "Mambalam", 
            "Arasankalani", "Arasankazhani"
        ],
        "Jagan / Karthik": [
            "Mylapore", "T Nagar", "T.Nagar"
        ],
        "Ventakesh / Dinikaran": [
            "Part Kottivakkam", "Semmancheri", "Semmanchery"
        ],
    }
    def normalize_text(text: str) -> str:
        if pd.isna(text) or text == "":
            return ""
        normalized = re.sub(r'[^\w\s]', '', str(text).strip().lower())
        return re.sub(r'\s+', ' ', normalized)
    
    def find_best_match(area_name: str) -> Optional[str]:
        if pd.isna(area_name) or area_name.strip() == "":
            return None
        normalized_area = normalize_text(area_name)
        for sales_person, areas in SALES_PERSON_AREAS.items():
            for mapped_area in areas:
                normalized_mapped = normalize_text(mapped_area)
                if normalized_area == normalized_mapped:
                    return sales_person
        return None
    
    def split_shared_assignments(df: pd.DataFrame, sales_col: str) -> pd.DataFrame:
        shared_mask = df[sales_col].str.contains('/', na=False)
        if not shared_mask.any():
            return df
        result_rows = []
        for idx, row in df.iterrows():
            sales_person = row[sales_col]
            if pd.isna(sales_person) or '/' not in sales_person:
                result_rows.append(row)
            else:
                salespeople = [sp.strip() for sp in sales_person.split('/')]
                result_rows.append({
                    'row': row,
                    'salespeople': salespeople,
                    'is_shared': True
                })
        final_rows = []
        shared_groups = {}
        for item in result_rows:
            if isinstance(item, dict) and item.get('is_shared'):
                key = ' / '.join(item['salespeople'])
                if key not in shared_groups:
                    shared_groups[key] = []
                shared_groups[key].append(item['row'])
            else:
                final_rows.append(item)
        for shared_key, rows in shared_groups.items():
            salespeople = shared_key.split(' / ')
            num_salespeople = len(salespeople)
            num_records = len(rows)
            if num_records == 1:
                row_copy = rows[0].copy()
                row_copy[sales_col] = salespeople[0]
                final_rows.append(row_copy)
            else:
                records_per_person = num_records // num_salespeople
                remainder = num_records % num_salespeople
                start_idx = 0
                for i, salesperson in enumerate(salespeople):
                    count = records_per_person + (1 if i < remainder else 0)
                    end_idx = start_idx + count
                    
                    for row in rows[start_idx:end_idx]:
                        row_copy = row.copy()
                        row_copy[sales_col] = salesperson
                        final_rows.append(row_copy)
                    start_idx = end_idx
        return pd.DataFrame(final_rows).reset_index(drop=True)
    try:
        if sheet_name:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(excel_file_path)
        if area_column_name not in df.columns:
            available_columns = list(df.columns)
            raise ValueError(f"Column '{area_column_name}' not found. Available columns: {available_columns}")
        result_df = df.copy()
        result_df[sales_person_column_name] = result_df[area_column_name].apply(find_best_match)
        result_df = split_shared_assignments(result_df, sales_person_column_name)
        matched_df = result_df[result_df[sales_person_column_name].notna()].copy()
        unmatched_df = result_df[result_df[sales_person_column_name].isna()].copy()
        matched_count = len(matched_df)
        unmatched_count = len(unmatched_df)        
        print(f"\n✅ Assignment completed:")
        print(f"  - Matched areas: {matched_count}")
        print(f"  - Unmatched areas: {unmatched_count}")
        
        if matched_count > 0:
            distribution = matched_df[sales_person_column_name].value_counts()
            for sp, count in distribution.items():
                print(f"  - {sp}: {count}")
        
        matched_temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        matched_df.to_excel(matched_temp_file.name, index=False)
        matched_file_path = matched_temp_file.name
        matched_temp_file.close()        
        unmatched_file_path = None
        if unmatched_count > 0:
            unmatched_areas = unmatched_df[area_column_name].unique()
            unmatched_temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            unmatched_df.to_excel(unmatched_temp_file.name, index=False)
            unmatched_file_path = unmatched_temp_file.name
            unmatched_temp_file.close()
            original_filename = os.path.basename(excel_file_path)
            alert_sent = send_unmatched_areas_alert(unmatched_df, original_filename)
            if alert_sent:
                print(f"✅ Alert email sent successfully to {os.getenv('RECIPIENT_MAIL')}")
            else:
                print("⚠️ Failed to send alert email")
        else:
            print("\n✅ All areas matched successfully! No unmatched records.")

        print('matched_file', matched_file_path,'\nunmatched_file',unmatched_file_path,)
        return matched_file_path        
    except Exception as e:
        print(f"❌ Error processing Excel file: {str(e)}")
        raise e
    
def send_records_alert(matched_df: pd.DataFrame, unmatched_df: pd.DataFrame, original_file_name: str = "input_file.xlsx") -> bool:
    try:
        sender_mailId = os.getenv("SENDER_MAIL", "riverpearlsolutions@gmail.com")
        passKey = os.getenv("APP_PASSWORD", "gwvcgbvjvttpvlja")
        recipient_email = os.getenv("RECIPIENT_MAIL")        
        if not sender_mailId or not passKey:
            print("Error: Email credentials not found")
            return False        
        if not recipient_email:
            print("Error: Recipient email not found")
            return False        
        if matched_df.empty and unmatched_df.empty:
            print("No records to report")
            return True        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")        
        temp_files = []        
        if not matched_df.empty:
            matched_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", mode='wb')
            matched_df.to_excel(matched_temp.name, index=False)
            matched_temp.close()
            temp_files.append({
                'path': matched_temp.name,
                'filename': f"Matched_Records_{timestamp}.xlsx",
                'type': 'matched'
            })        
        if not unmatched_df.empty:
            unmatched_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", mode='wb')
            unmatched_df.to_excel(unmatched_temp.name, index=False)
            unmatched_temp.close()
            temp_files.append({
                'path': unmatched_temp.name,
                'filename': f"Unmatched_Records_{timestamp}.xlsx",
                'type': 'unmatched'
            })        
        msg = MIMEMultipart()
        msg['From'] = sender_mailId
        msg['To'] = recipient_email
        msg['Subject'] = f"Records Report: Matched & Unmatched - {original_file_name}"
        total_matched = len(matched_df)
        total_unmatched = len(unmatched_df)
        total_records = total_matched + total_unmatched        
        body = f'''
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <div style="max-width: 650px; margin: 0 auto; padding: 20px;">
                <div style="background-color: #4a90e2; color: white; padding: 15px; border-radius: 5px;">
                    <h2 style="margin: 0;">📊 Records Processing Report</h2>
                </div>
                
                <div style="padding: 20px; background-color: #f9f9f9; margin-top: 20px; border-radius: 5px;">
                    <p>Dear Team,</p>
                    <p>The system has completed processing <strong>{original_file_name}</strong>. Please find the summary and attached files below.</p>
                    
                    <div style="background-color: white; padding: 15px; border-left: 4px solid #4a90e2; margin: 20px 0;">
                        <h3 style="margin-top: 0; color: #4a90e2;">Processing Summary</h3>
                        <ul style="list-style: none; padding: 0;">
                            <li>📁 <strong>Source File:</strong> {original_file_name}</li>
                            <li>📅 <strong>Generated On:</strong> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</li>
                            <li>📈 <strong>Total Records:</strong> {total_records}</li>
                        </ul>
                    </div>
                    
                    <div style="display: flex; gap: 15px; margin: 20px 0;">
                        <div style="flex: 1; background-color: #d4edda; padding: 15px; border-radius: 5px; border-left: 4px solid #28a745;">
                            <h4 style="margin-top: 0; color: #155724;">✅ Matched Records</h4>
                            <p style="font-size: 24px; font-weight: bold; margin: 10px 0; color: #155724;">{total_matched}</p>
                            <p style="font-size: 12px; color: #155724; margin: 0;">
                                {f'{(total_matched/total_records*100):.1f}%' if total_records > 0 else '0%'} of total
                            </p>
                        </div>
                        
                        <div style="flex: 1; background-color: #f8d7da; padding: 15px; border-radius: 5px; border-left: 4px solid #dc3545;">
                            <h4 style="margin-top: 0; color: #721c24;">⚠️ Unmatched Records</h4>
                            <p style="font-size: 24px; font-weight: bold; margin: 10px 0; color: #721c24;">{total_unmatched}</p>
                            <p style="font-size: 12px; color: #721c24; margin: 0;">
                                {f'{(total_unmatched/total_records*100):.1f}%' if total_records > 0 else '0%'} of total
                            </p>
                        </div>
                    </div>
                    
                    <div style="background-color: #fff3cd; padding: 15px; border-radius: 5px; margin: 20px 0;">
                        <h4 style="margin-top: 0; color: #856404;">📋 Matching Criteria:</h4>
                        <p style="margin: 5px 0;">Records are considered <strong>matched</strong> if:</p>
                        <ol style="margin: 10px 0;">
                            <li><strong>Dwelling Unit Info</strong> is not empty/null, OR</li>
                            <li><strong>Nature of Development</strong> contains keywords: school building, hospital, college, inst, kalayaan mandapam</li>
                        </ol>
                        <p style="color: #856404; font-style: italic;">Unmatched records do not meet either of these conditions.</p>
                    </div>
                    
                    <div style="background-color: #e7f3ff; padding: 15px; border-radius: 5px; margin: 20px 0;">
                        <h4 style="margin-top: 0; color: #004085;">📎 Attachments:</h4>
                        <ul style="margin: 5px 0;">
                            {'<li>✅ <strong>Matched_Records.xlsx</strong> - Contains all matched records</li>' if total_matched > 0 else ''}
                            {'<li>⚠️ <strong>Unmatched_Records.xlsx</strong> - Contains all unmatched records</li>' if total_unmatched > 0 else ''}
                        </ul>
                    </div>
                    
                    <p><strong>Action Required:</strong></p>
                    <ol>
                        <li>Review both attached Excel files</li>
                        <li>For unmatched records, verify if "Dwelling Unit Info" should be populated</li>
                        <li>Check if "Nature of Development" should contain relevant keywords</li>
                        <li>Update the records and reprocess the file if needed</li>
                    </ol>
                </div>
                
                <hr style="border: none; border-top: 1px solid #ddd; margin: 30px 0;">
                
                <div style="font-size: 12px; color: #666;">
                    <p><strong>VPEARL SOLUTIONS - An AI Company, Chennai</strong></p>
                    <p>
                        🌐 <a href="https://vpearlsolutions.com/" target="_blank" style="color: #4a90e2;">Website</a> | 
                        🔗 <a href="https://www.linkedin.com/company/vpealsoutions/" target="_blank" style="color: #4a90e2;">LinkedIn</a> | 
                        📷 <a href="https://www.instagram.com/vpearl_solutions" target="_blank" style="color: #4a90e2;">Instagram</a> | 
                        📘 <a href="https://www.facebook.com/profile.php?id=61572978223085" target="_blank" style="color: #4a90e2;">Facebook</a>
                    </p>
                    <p style="font-size: 10px; color: #999;">
                        <em>This is an automated report. Please do not reply to this email.</em>
                    </p>
                </div>
            </div>
        </body>
        </html>
        '''
        
        msg.attach(MIMEText(body, 'html'))        
        for file_info in temp_files:
            if os.path.exists(file_info['path']):
                with open(file_info['path'], 'rb') as f:
                    attachment = MIMEApplication(f.read(), _subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    attachment.add_header('Content-Disposition', 'attachment', filename=file_info['filename'])
                    msg.attach(attachment)
            else:
                print(f"Warning: File not found at {file_info['path']}")        
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_mailId, passKey)
            server.sendmail(sender_mailId, recipient_email, msg.as_string())
        print(f"✅ Email report sent successfully to {recipient_email}")
        print(f"   - Matched records: {total_matched}")
        print(f"   - Unmatched records: {total_unmatched}")        
        for file_info in temp_files:
            try:
                os.unlink(file_info['path'])
            except:
                pass
        return True
    except Exception as e:
        print(f"❌ Error in send_records_alert function: {str(e)}")
        return False

def separate_and_store_temp(filepath, send_email=True):
    keywords = ["premium fsi","units","mall","theatre building","screens","dwelling units","dwellings","school building", "hospital", "college", "inst", "kalyana mandapam","auditorium","service apartment","service apartments"]
    try:
        df = pd.read_excel(filepath)
        original_file_name = os.path.basename(filepath)
        required_cols = ["Dwelling Unit Info", "Nature of Development"]
        for col in required_cols:
            if col not in df.columns:
                raise ValueError(f"Missing required column: {col}")        
        cond1 = df["Dwelling Unit Info"].notna() & (df["Dwelling Unit Info"].astype(str).str.strip() != "")                
        nature_lower = df["Nature of Development"].astype(str).str.lower().str.strip()
        cond2 = df["Dwelling Unit Info"].isna() | (df["Dwelling Unit Info"].astype(str).str.strip() == "")
        cond2 = cond2 & nature_lower.apply(lambda x: any(k in x for k in keywords))
        matched_df = df[cond1 | cond2]
        unmatched_df = df[~(cond1 | cond2)]        
        matched_temp_file = tempfile.NamedTemporaryFile(delete=False, suffix="_matched.xlsx")
        matched_df.to_excel(matched_temp_file.name, index=False)
        print(f"✅ Matched data saved to: {matched_temp_file.name}")
        print(f"   Total matched records: {len(matched_df)}")        
        unmatched_temp_file = tempfile.NamedTemporaryFile(delete=False, suffix="_unmatched.xlsx")
        unmatched_df.to_excel(unmatched_temp_file.name, index=False)
        print(f"✅ Unmatched data saved to: {unmatched_temp_file.name}")
        print(f"   Total unmatched records: {len(unmatched_df)}")        
        if send_email:
            print("\n📧 Sending email report with matched and unmatched records...")
            email_sent = send_records_alert(matched_df, unmatched_df, original_file_name)
            if email_sent:
                print("✅ Email report sent successfully!")
            else:
                print("⚠️ Failed to send email report")
        return matched_temp_file.name
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

# Zoho CRM Configuration
CLIENT_ID = "1000.FG0RKGI29XWAGYJFGN8BAPYLK298PR"
CLIENT_SECRET = "4943e92f55611b94772c38a65acf466708a91b4da3"
REDIRECT_URL = "https://www.google.com/"
ORG_ID = "60037625008"
EMAIL_ADDRESS = "abhishek.rg.abppl@gmail.com"
PASSWORD = "Ajantha@2025"
AUTH_URL = "https://accounts.zoho.com/oauth/v2/auth"
TOKEN_URL = "https://accounts.zoho.in/oauth/v2/token"
API_BASE_URL = "https://www.zohoapis.in/crm/v2"
TOKEN_FILE_NAME = "zoho_tokens.json"
ZOHO_MODEL_NAME = "CMDA"