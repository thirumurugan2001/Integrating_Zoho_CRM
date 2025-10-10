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

def separate_and_store_temp(filepath):
    keywords = ["school building", "hospital", "college", "inst", "kalayaan mandapam"]
    try:
        df = pd.read_excel(filepath)        
        required_cols = ["Dwelling Unit Info", "Nature of Development"]
        for col in required_cols:
            if col not in df.columns:
                raise ValueError(f"Missing required column: {col}")        
        cond1 = df["Dwelling Unit Info"].notna() & (df["Dwelling Unit Info"].astype(str).str.strip() != "")        
        cond2 = df["Dwelling Unit Info"].isna() | (df["Dwelling Unit Info"].astype(str).str.strip() == "")
        cond2 = cond2 & df["Nature of Development"].astype(str).str.lower().apply(
            lambda x: any(k in x for k in keywords))        
        filtered_df = df[cond1 | cond2]        
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        filtered_df.to_excel(temp_file.name, index=False)
        print(f"✅ Filtered data saved to: {temp_file.name}")
        return temp_file.name
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

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
        # Adambakkam
        "Abhishek": ["Alandur", "Alandur Guindy", "Guindy", "Madipakkam", 
            "Medavakkam", "Nanganallur", "Pallikaranai", "Thalakananchery", 
            "Thalakkanancheri", "Thalakkananchery", "Thalakkancheri", "Velachery"
        ],
        "Jagan": [
            "Adyar", "Athipattu", "Egmore", "Kottur", "Koyambedu", "koyambedu", 
            "Koyembedu", "Mogappair", "Mullam", "Naduvakarai", "Naduvankarai", 
            "Naduvankkarai", "Nekundram", "Nerkundram", "Nolambur", "Nungambakkam", 
            "Pallipattu", "Part of Thirumangalam", "Periyakudal", 
            "Secretariat Colony Kilpauk Chennai.", "Urur", "Vada Agaram", "Vepery"
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
            "Ward No.C of Tambaram", "Ward No.D of Tambaram"
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
            "Vilangadupakkam", "Villivakkam", "Ward No. I of Paruthipattu"
        ],
        "Karthik / Ventakesh": [
            "Gerugambakkam", "Kollacheri", "Kulappakkam", "Kuthambakkam", 
            "Poonamallee", "Rendamkattalai", "Rendankattalai", "Sikkarayapuram", 
            "Vellavedu", "Zamin Pallavaram", "Zamin Pallvaram", "Mambalam", 
            "Arasankalani", "Arasankazhani"
        ],
        "Jagan / Karthik": [
            "Mylapore", "T. Nagar", "T.Nagar"
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