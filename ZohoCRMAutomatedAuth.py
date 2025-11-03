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
                return False
    
    def handle_tfa_banner_page(self, driver):
        current_url = driver.current_url
        if "tfa-banner" in current_url or "announcement" in current_url:
            continue_selectors = [("xpath", "//button[contains(text(), 'Continue')]"),("xpath", "//button[contains(text(), 'Skip')]"),("xpath", "//button[contains(text(), 'Later')]"),("xpath", "//button[contains(text(), 'Not now')]"),("xpath", "//a[contains(text(), 'Continue')]"),("xpath", "//a[contains(text(), 'Skip')]"),("xpath", "//a[contains(text(), 'Later')]"),("xpath", "//input[@value='Continue']"),("xpath", "//input[@value='Skip']"),("id", "continue"),("id", "skip"),("id", "later"), ("css", ".continue-btn"),("css", ".skip-btn"),("css", "button.primary"),("css", "a.primary"),("xpath", "//button[contains(@class, 'primary')]"),("xpath", "//a[contains(@class, 'primary')]")]
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
                    time.sleep(10)
                    return True
        return False
    
    def get_authorization_url(self):
        params = {'scope': 'ZohoCRM.modules.ALL,ZohoCRM.settings.ALL,ZohoCRM.users.ALL','client_id': self.client_id,'response_type': 'code','access_type': 'offline','redirect_uri': self.redirect_uri}
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
            email_selectors = [("id", "login_id"),("name", "LOGIN_ID"),("xpath", "//input[@type='email']"),("xpath", "//input[@placeholder='Email ID']"),("css", "input[type='email']"),("xpath", "//input[contains(@class, 'email')]")]
            email_element, _, _ = self.wait_and_find_element(driver, email_selectors, 30)
            if not email_element:
                self.debug_page(driver)
                return False            
            if not self.safe_send_keys(driver, email_element, self.email, "email field"):
                return False
            next_selectors = [("id", "nextbtn"),("id", "signin_submit"),("xpath", "//button[contains(text(), 'Next')]"),("xpath", "//button[contains(text(), 'Continue')]"),("xpath", "//input[@value='Next']"),("xpath", "//input[@value='Continue']"),("css", "button[type='submit']"),("xpath", "//button[@type='submit']")]
            next_element, _, _ = self.wait_and_find_element(driver, next_selectors, 20)
            if next_element:
                if self.safe_click(driver, next_element, "next button"):
                    time.sleep(3)           
            password_selectors = [("id", "password"),("name", "PASSWORD"),("xpath", "//input[@type='password']"),("css", "input[type='password']"),("xpath", "//input[contains(@class, 'password')]")]
            password_element, _, _ = self.wait_and_find_element(driver, password_selectors, 30)
            if not password_element:
                self.debug_page(driver)
                return False
            if not self.safe_send_keys(driver, password_element, self.password, "password field"):
                return False            
            signin_selectors = [("id", "nextbtn"),("id", "signin_submit"),("xpath", "//button[contains(text(), 'Sign in')]"),("xpath", "//button[contains(text(), 'Sign In')]"),("xpath", "//button[contains(text(), 'Login')]"),("xpath", "//input[@value='Sign in']"),("xpath", "//input[@value='Sign In']"),("css", "button[type='submit']"),("xpath", "//button[@type='submit']")]
            signin_element, _, _ = self.wait_and_find_element(driver, signin_selectors, 20)
            if not signin_element:
                self.debug_page(driver)
                return False
            if not self.safe_click(driver, signin_element, "sign in button"):
                return False
            time.sleep(5)            
            if self.handle_tfa_banner_page(driver):
                time.sleep(3)            
            accept_selectors = [("xpath", "//button[contains(text(), 'Accept')]"),("xpath", "//button[contains(text(), 'Allow')]"),("xpath", "//button[contains(text(), 'Authorize')]"),("xpath", "//input[@value='Accept']"),("xpath", "//input[@value='Allow']"),("id", "accept"),("id", "allow"),("css", "button.accept"),("css", "button.allow"),("xpath", "//button[contains(@class, 'accept')]"),("xpath", "//button[contains(@class, 'allow')]")]
            accept_element, _, _ = self.wait_and_find_element(driver, accept_selectors, 15)
            if accept_element:
                if self.safe_click(driver, accept_element, "accept button"):
                    time.sleep(3)
            else:
                pass
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
                        info = {'tag': tag,'id': element.get_attribute('id'),'name': element.get_attribute('name'),'type': element.get_attribute('type'),'class': element.get_attribute('class'),'text': element.text[:50] if element.text else '','displayed': element.is_displayed(),'enabled': element.is_enabled()}
                        elements_info.append(info)
                    except:
                        continue
        except Exception as e:
            pass
    
    def get_access_token(self, authorization_code):        
        data = {'grant_type': 'authorization_code','client_id': self.client_id,'client_secret': self.client_secret,'redirect_uri': self.redirect_uri,'code': authorization_code}
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
                return False
        except Exception as e:
            return False
    
    def refresh_access_token(self):
        if not self.refresh_token:
            return False
        data = {'grant_type': 'refresh_token','client_id': self.client_id,'client_secret': self.client_secret,'refresh_token': self.refresh_token}
        try:
            response = requests.post(self.token_url, data=data)
            if response.status_code == 200:
                token_data = response.json()
                self.access_token = token_data['access_token']
                expires_in = token_data.get('expires_in', 3600)
                self.token_expires_at = datetime.now() + timedelta(seconds=expires_in)
                self.save_tokens()
                return True
            else:
                return False
        except Exception as e:
            return False
    
    def save_tokens(self):
        token_data = {'access_token': self.access_token,'refresh_token': self.refresh_token,'expires_at': self.token_expires_at.isoformat() if self.token_expires_at else None,'client_id': self.client_id}
        try:
            with open(self.token_file, 'w') as f:
                json.dump(token_data, f, indent=2)
        except Exception as e:
            pass

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

    def get_user_id_by_name(self, sales_person_name):
        if not self.ensure_valid_token():
            return None
        user_mapping = {
            "Abhishek": "60037625008",
            "Jagan": "60037625008",
            "Karthik": "60037625008",
            "Ventakesh": "60037625008",
            "Dinikaran": "60037625008",
            "Balachander": "60037625008",
        }
        return user_mapping.get(sales_person_name)

    def format_record_for_zoho(self, record):
        formatted_record = {}
        field_mapping = {
            "Sales Person": "Lead_Owner", 
            "Email ID": "Email",
            "Mobile No.": "Mobile_Number",
            "Date of permit": "Date_of_Permit",
            "Applicant Name": "Lead_Name",
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
        if record.get("Applicant Name") and not pd.isna(record.get("Applicant Name")):
            formatted_record["Name"] = str(record["Applicant Name"]).strip()
        elif record.get("Company_Name") and not pd.isna(record.get("Company_Name")):
            formatted_record["Name"] = str(record["Company_Name"]).strip()
        else:
            formatted_record["Name"] = f"Record_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        if "Sales Person" in record and record["Sales Person"] and not pd.isna(record["Sales Person"]):
            sales_person = str(record["Sales Person"]).strip()
            user_id = self.get_user_id_by_name(sales_person)
            if user_id:
                formatted_record["Lead_Owner"] = user_id
            else:
                pass
        if "No_of_bathrooms" in formatted_record:
            pass
        else:
            pass        
        formatted_record['Lead_Source'] = "Digital Leads"
        return formatted_record

    def push_records_to_zoho(self, records, batch_size=100):
        if not self.ensure_valid_token():
            return False
        if not records:
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
            headers = {'Authorization': f'Zoho-oauthtoken {self.access_token}','Content-Type': 'application/json'}
            payload = {'data': formatted_batch,'trigger': ['approval', 'workflow', 'blueprint']}
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
                                print(f"✅ Record created successfully: {result.get('message', 'Success')}")
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
                traceback.print_exc()
                failed_records += len(formatted_batch)        
        print(f"\n✅ Push completed: {successful_records} successful, {failed_records} failed out of {total_records} total")
        return successful_records > 0

    def test_api_connection(self):
        if not self.ensure_valid_token():
            return False
        url = f"{self.api_base_url}/settings/modules"
        headers = {'Authorization': f'Zoho-oauthtoken {self.access_token}','Content-Type': 'application/json'}
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                modules = response.json()
                module_names = [module['api_name'] for module in modules.get('modules', [])]
                print(f"Available modules: {module_names}")
                if self.zoho_model_name in module_names:
                    return True
                else:
                    return False
            else:
                return False
        except Exception as e:
            return False

    def get_module_fields(self):
        if not self.ensure_valid_token():
            return False
        url = f"{self.api_base_url}/settings/fields?module={self.zoho_model_name}"
        headers = {'Authorization': f'Zoho-oauthtoken {self.access_token}','Content-Type': 'application/json'}
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                return True
            else:
                return False
        except Exception as e:
            return False
    
    def create_lead_from_cmda_record(self, cmda_record):        
        if not self.ensure_valid_token():
            return False        
        lead_data = {
            "Company": self.clean_value(cmda_record.get("Company_Name", "")),
            "Last_Name": self.clean_value(cmda_record.get("Applicant Name", "Unknown Applicant")), 
            "Email": self.clean_value(cmda_record.get("Email ID", "")),
            "Phone": self.clean_value(cmda_record.get("Mobile No.", "")),
            "First_Name": self.extract_first_name(cmda_record.get("Applicant Name", "")),
            "Billing_Area": self.clean_value(cmda_record.get("Area Name", "")),
            "Street": self.clean_value(cmda_record.get("Applicant Address", cmda_record.get("Site Address", ""))),
            "City": "Chennai", 
            "State": "Tamil Nadu", 
            "Country": "India",            
            "Project_Name": self.clean_value(cmda_record.get("Company_Name", "")),
            "How_Much_Square_Feet": self.clean_value(cmda_record.get("How_Much_Square_Feet", "")),
            "Site_Area": self.clean_value(cmda_record.get("Site_Area", "")),
            "Reference": self.clean_value(cmda_record.get("Reference", "")),
            "Architect_Name": self.clean_value(cmda_record.get("Architect Name", "")),            
            "Description": self.build_description(cmda_record),
            "Lead_Source": "Digital Leads",
            "Future_Projects": self.clean_value(cmda_record.get("Future_Projects", "-None-")),            
            "Planning_Permission_No": self.clean_value(cmda_record.get("Planning Permission No.", "")),
            "Permit_No": self.clean_value(cmda_record.get("Permit No.", "")),
            "File_No": self.clean_value(cmda_record.get("File No.", "")),
        }        
        self.handle_numeric_fields(lead_data, cmda_record)        
        self.handle_picklist_fields(lead_data, cmda_record)        
        self.handle_date_fields(lead_data, cmda_record)        
        self.handle_sales_person_assignment(lead_data, cmda_record)        
        lead_data = self.final_data_cleaning(lead_data)        
        url = f"{self.api_base_url}/Leads"
        headers = {
            'Authorization': f'Zoho-oauthtoken {self.access_token}',
            'Content-Type': 'application/json'
        }
        payload = {
            'data': [lead_data],
            'trigger': ['workflow']
        }       
        try:
            response = requests.post(url, json=payload, headers=headers)
            if response.status_code == 201:
                result = response.json()
                print("✅ Lead created successfully in Zoho CRM!")
                if 'data' in result:
                    for item in result['data']:
                        if item.get('status') == 'success':
                            lead_id = item.get('details', {}).get('id', 'Unknown')
                            print(f"🎉 Lead Created Successfully! ID: {lead_id}")
                            print(f"📋 Complete Response: {item}")
                        else:
                            error_msg = item.get('message', 'Unknown error')
                            error_details = item.get('details', 'No details')
                            print(f"❌ Lead creation failed: {error_msg}")
                            print(f"🔍 Error Details: {error_details}")
                return True
            else:
                print(f"❌ Failed to create Lead. Status: {response.status_code}")
                print(f"🔍 Response: {response.text}")
                return False                
        except Exception as e:
            print(f"❌ Error creating Lead in Zoho CRM: {e}")
            traceback.print_exc()
            return False

    def clean_value(self, value):
        if value is None or pd.isna(value):
            return ""        
        value_str = str(value).strip()
        if value_str.lower() in ['', 'nan', 'none', 'null']:
            return ""        
        return value_str

    def extract_first_name(self, applicant_name):
        if not applicant_name or pd.isna(applicant_name):
            return ""        
        name_parts = str(applicant_name).split()
        if name_parts:
            for part in name_parts:
                if part.istitle() and len(part) > 2 and not part.isupper():
                    return part            
            return name_parts[0]        
        return ""

    def build_description(self, cmda_record):
        description_parts = []        
        nature = cmda_record.get("Nature of Development", "")
        if nature and not pd.isna(nature):
            description_parts.append(f"Nature: {nature}")        
        site_address = cmda_record.get("Site Address", "")
        if site_address and not pd.isna(site_address):
            description_parts.append(f"Site: {site_address}")        
        applicant_address = cmda_record.get("Applicant Address", "")
        if applicant_address and not pd.isna(applicant_address):
            description_parts.append(f"Applicant Address: {applicant_address}")        
        file_no = cmda_record.get("File No.", "")
        permit_no = cmda_record.get("Permit No.", "")
        planning_no = cmda_record.get("Planning Permission No.", "")        
        if file_no and not pd.isna(file_no):
            description_parts.append(f"File: {file_no}")
        if permit_no and not pd.isna(permit_no):
            description_parts.append(f"Permit: {permit_no}")
        if planning_no and not pd.isna(planning_no):
            description_parts.append(f"Planning: {planning_no}")        
        app_date = cmda_record.get("Date of Application", "")
        if app_date and not pd.isna(app_date):
            description_parts.append(f"Application Date: {app_date}")        
        return " | ".join(description_parts) if description_parts else "CMDA Record"

    def handle_numeric_fields(self, lead_data, cmda_record):
        try:
            dwelling_units = cmda_record.get("Dwelling Unit Info", "0")
            if dwelling_units and self.clean_value(dwelling_units):
                numbers = re.findall(r'\d+', str(dwelling_units))
                if numbers:
                    dwelling_value = int(numbers[0])
                    lead_data["No_of_Bathrooms"] = dwelling_value * 2 
                else:
                    lead_data["No_of_Bathrooms"] = 0
            else:
                lead_data["No_of_Bathrooms"] = 0
        except (ValueError, TypeError) as e:
            lead_data["No_of_Bathrooms"] = 0
        
        try:
            dwelling_units = cmda_record.get("Dwelling Unit Info", "0")
            if dwelling_units and self.clean_value(dwelling_units):
                numbers = re.findall(r'\d+', str(dwelling_units))
                if numbers:
                    lead_data["No_of_Units"] = int(numbers[0])
                else:
                    lead_data["No_of_Units"] = 0
            else:
                lead_data["No_of_Units"] = 0
        except (ValueError, TypeError) as e:
            lead_data["No_of_Units"] = 0

    def handle_picklist_fields(self, lead_data, cmda_record):
        which_brand = cmda_record.get("Which_Brand_Looking_for", "")
        if which_brand and self.clean_value(which_brand):
            lead_data["Which_Brand_Looking_for"] = self.clean_value(which_brand)        
        if not lead_data.get("Future_Projects"):
            lead_data["Future_Projects"] = "-None-"
        if not lead_data.get("Lead_Source"):
            lead_data["Lead_Source"] = "Digital Leads"

    def handle_date_fields(self, lead_data, cmda_record):
        date_of_permit = cmda_record.get("Date of permit", "")
        if date_of_permit and self.clean_value(date_of_permit):
            try:
                if isinstance(date_of_permit, str) and len(date_of_permit) == 10 and '-' in date_of_permit:
                    dt = datetime.strptime(date_of_permit, "%d-%m-%Y")
                    lead_data["Date_of_Permit"] = dt.strftime("%Y-%m-%d")
                else:
                    lead_data["Date_of_Permit"] = self.clean_value(date_of_permit)
            except ValueError:
                lead_data["Date_of_Permit"] = self.clean_value(date_of_permit)
        date_of_application = cmda_record.get("Date of Application", "")
        if date_of_application and self.clean_value(date_of_application):
            try:
                if isinstance(date_of_application, str) and '/' in date_of_application:
                    dt = datetime.strptime(date_of_application, "%d/%m/%Y")
                    lead_data["Date_of_Application"] = dt.strftime("%Y-%m-%d")
                else:
                    lead_data["Date_of_Application"] = self.clean_value(date_of_application)
            except ValueError:
                lead_data["Date_of_Application"] = self.clean_value(date_of_application)

    def handle_sales_person_assignment(self, lead_data, cmda_record):
        sales_person = cmda_record.get("Sales Person", "")
        if sales_person and self.clean_value(sales_person):
            sales_person_clean = sales_person.strip()            
            owner_field_mapping = {
                "Abhishek": "Ameen_Syed_Owne",
                "Balachander": "Balachander_Owner",
                "Dinikaran": "Dinakaran_Owner",
                "Jagan": "Jagan_Owner",
                "Karthik": "Karthik_Owner",
                "Venkatesh": "Venkatesh_Owner",
                "Ramanunjam": "Ramanunjam_Owner"
            }            
            owner_field = owner_field_mapping.get(sales_person_clean)
            if owner_field:
                lead_data[owner_field] = "Yes"
                print(f"✅ Assigned {sales_person_clean} to field: {owner_field}")
            else:
                print(f"⚠️  No owner field mapping found for: {sales_person_clean}")

    def final_data_cleaning(self, lead_data):
        cleaned_data = {}        
        for key, value in lead_data.items():
            if value is None:
                continue
            if isinstance(value, str) and value.strip() == "":
                continue
            if pd.isna(value):
                continue            
            cleaned_data[key] = value        
        return cleaned_data
   