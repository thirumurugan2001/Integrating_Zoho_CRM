import pandas as pd
import re
from fuzzywuzzy import fuzz
from typing import Optional
import tempfile
import os
   
def assign_sales_person_to_areas(excel_file_path: str,area_column_name: str = 'Area Name',sales_person_column_name: str = 'Sales Person',sheet_name: str = None,fuzzy_match_threshold: int = 85
) -> str:
    SALES_PERSON_AREAS = {
        "Abhishek": [
            "Adambakkam", "Alandur", "Alandur Guindy", "Guindy", "Madipakkam", 
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
        "Jagan / Balacahnder": [
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
        "Ventakesh / Karthik": [
            "Arasankalani", "Arasankazhani"
        ],
        "Karthik / Ventakesh": [
            "Gerugambakkam", "Kollacheri", "Kulappakkam", "Kuthambakkam", 
            "Poonamallee", "Rendamkattalai", "Rendankattalai", "Sikkarayapuram", 
            "Vellavedu", "Zamin Pallavaram", "Zamin Pallvaram"
        ],
        "Karthik / Jagan": [
            "Mambalam"
        ],
        "Jagan / karthik": [
            "Mylapore", "T. Nagar", "T.Nagar"
        ],
        "Ventakesh / Dinikaran": [
            "Part Kottivakkam", "Semmancheri", "Semmanchery"
        ],
        "Jagan /  Balachander": [
            "Sholavaram", "Sirugavoor"
        ]
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
        best_match = None
        best_score = 0
        for sales_person, areas in SALES_PERSON_AREAS.items():
            for mapped_area in areas:
                normalized_mapped = normalize_text(mapped_area)
                if normalized_area == normalized_mapped:
                    return sales_person
                score = fuzz.ratio(normalized_area, normalized_mapped)
                if score >= fuzzy_match_threshold and score > best_score:
                    best_match = sales_person
                    best_score = score
        return best_match
    try:
        if sheet_name:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(excel_file_path)
        if area_column_name not in df.columns:
            available_columns = list(df.columns)
            raise ValueError(f"Column '{area_column_name}' not found. Available columns: {available_columns}")
        result_df = df.copy()
        print("Assigning sales persons based on area matching...")
        result_df[sales_person_column_name] = result_df[area_column_name].apply(find_best_match)
        matched_count = result_df[sales_person_column_name].notna().sum()
        unmatched_count = len(result_df) - matched_count
        print(f"Assignment completed:\n  - Matched areas: {matched_count}\n  - Unmatched areas: {unmatched_count}")

        if unmatched_count > 0:
            unmatched_areas = result_df[result_df[sales_person_column_name].isna()][area_column_name].unique()
            print(f"Unmatched areas: {list(unmatched_areas)}")

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        result_df.to_excel(temp_file.name, index=False)
        temp_file_path = temp_file.name
        temp_file.close()

        print(f"Results saved to temporary file: {temp_file_path}")
        return temp_file_path

    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        raise


def example_usage():
    try:
        temp_path = assign_sales_person_to_areas(excel_file_path="Testing_data.xlsx")
        print(f"Basic assignment completed successfully! Temp file: {temp_path}")
    except Exception as e:
        print(f"Error in basic usage: {e}")


if __name__ == "__main__":
    example_usage()
