from ZohoCRMAutomatedAuth import ZohoCRMAutomatedAuth
from helper import excel_to_json

def lead_import(file_path):
    try : 
        crm = ZohoCRMAutomatedAuth()    
        if crm.test_api_connection():
            crm.get_module_fields()        
            records = excel_to_json(file_path)       
            if records:
                success = crm.push_records_to_zoho(records)
                if success:
                    return {
                        "message": "Records pushed successfully!",
                        "statusCode": 200,
                        "status": True,
                    }
                else:
                    return {
                        "message": "Failed to push some or all records",
                        "statusCode": 400,  
                        "status": False,
                    }
            else:
                return {
                    "message": "No records found in Excel file",
                    "statusCode": 400,  
                    "status": False,
                }
        else:
            return {
                    "message": "API connection failed!",
                    "statusCode": 400,  
                    "status": False,
                    "data": [{}]
                }
    except Exception as e:
        return {
            "message": str(e),
            "statusCode": 400,
            "status": False,
            "data": [{}]
        }
        
data = [
  {
    "Abhishek": [
      "Adambakkam",
      "Alandur",
      "Alandur Guindy",
      "Guindy",
      "Madipakkam",
      "Medavakkam",
      "Nanganallur",
      "Pallikaranai",
      "Thalakananchery",
      "Thalakkanancheri",
      "Thalakkananchery",
      "Thalakkancheri",
      "Velachery"
    ]
  },
  {
    "Jagan": [
      "Adyar",
      "Athipattu",
      "Egmore",
      "Kottur",
      "Koyambedu",
      "koyambedu",
      "Koyembedu",
      "Mogappair",
      "Mullam",
      "Naduvakarai",
      "Naduvankarai",
      "Naduvankkarai",
      "Nekundram",
      "Nerkundram",
      "Nolambur",
      "Nungambakkam",
      "Pallipattu",
      "Part of Thirumangalam",
      "Periyakudal",
      "Secretariat Colony Kilpauk Chennai.",
      "Urur",
      "Vada Agaram",
      "Vepery"
    ]
  },
  {
    "Karthik": [
      "Arumbakkam",
      "Ayyappanthangal",
      "Ekkaduthangal",
      "Goparasanallur",
      "Kalikundram",
      "Kanagam",
      "Karambakkam",
      "Kodambakkam",
      "Kolapakkam",
      "Kulamanivakkam",
      "Madhananthapuram",
      "Madhandhapuram",
      "Manapakkam",
      "Mangadu-B",
      "Moulivakkam",
      "Noombal",
      "Pammal",
      "Panaveduthottam",
      "Parivakkam",
      "Porur",
      "Puliyur",
      "Saligramam",
      "Tharapakkam",
      "Valasaravakkam",
      "Virugambakkam",
      "Voyalanallur-A"
    ]
  },
  {
    "Ventakesh": [
      "Agaramthen",
      "Anakaputhur",
      "Chembarambakkam",
      "Cowl Bazaar",
      "Gowrivakkam",
      "Karapakkam",
      "Kaspapuram",
      "Kulathuvancheri",
      "Kundrathur",
      "Kundrathur - A",
      "Kundrathur - B",
      "Kundrathur-A",
      "Kundrathur-B",
      "Malayambakkam",
      "Manancheri",
      "Mannivakkam",
      "Meppedu",
      "Mudichur",
      "Mullam",
      "Nandambakkam",
      "Nanmangalam",
      "Naduveerapattu",
      "Nedungundram",
      "Nedunkundram",
      "Nemilichery",
      "Ottiyambakkam",
      "Palanthandalam",
      "Pallavaram",
      "Pallavarm",
      "Perumbakkam",
      "Perungalathur",
      "Rajakilpakkam",
      "S.Kulathur",
      "Selaiyur",
      "Sirukalathur",
      "Tambaram",
      "Thirumudivakkam",
      "Thiruneermalai",
      "Thiruvancheri",
      "Vandalur",
      "Varadarajapuram",
      "Varadharajapuram",
      "Vengaivasal",
      "Vengambakkam",
      "Ward No.C of Tambaram",
      "Ward No.D of Tambaram"
    ]
  },
  {
    "Dinikaran": [
      "Kottivakkam",
      "Kovilambakkam",
      "Neelangarai",
      "Okkiam Thoraipakkam",
      "Okkiyam Thoraipakkam",
      "part of Sholinganallur",
      "Perungudi",
      "Sholinganallur",
      "Thiiruvanmiyur",
      "Thiruvanmiyur",
      "Thoraipakkam"
    ]
  },
  {
    "Balachander": [
      "Agraharammel",
      "Angadu",
      "Layon Pullion",
      "Maduravoyal"
    ]
  },
  {
    "Sithalapakkam": [
      "Sithalapakkam"
    ]
  },
  {
    "Jagan / Balacahnder": [
      "Adayalampattu",
      "Alamathi",
      "Ambathur",
      "Ambattur",
      "Arumandai",
      "at Kondakarai Kuruvimedu Panchayat Road and",
      "at Orakkadu",
      "at Puzhal",
      "Ayanambakkam",
      "Ayanavaram",
      "Budur",
      "BUDUR",
      "Chintadripet",
      "Girudalapuram",
      "Kannapalayam",
      "Karanodai",
      "Karunakaracheri",
      "Kathirvedu",
      "Korattur",
      "Korattur A",
      "Kosapur",
      "Kovilpadagai",
      "Layon Grant",
      "Madhavaram",
      "Mijur",
      "Minjur",
      "Minjur II",
      "Nayar-II",
      "Nemam",
      "Oragadam",
      "Orakkadu",
      "Padi",
      "Padiyanallur",
      "Pakkam",
      "Palanjur",
      "Paleripattu",
      "part of Ayapakkam",
      "Paruthipattu",
      "Perambur",
      "Peravallur",
      "Periyamullaivoyal",
      "Perungavur",
      "Peruvallur",
      "Ponneri",
      "Purasaiwalkam",
      "Purasalwalkam",
      "Purursawalkkam",
      "Purusawalkam",
      "Seemapuram",
      "Sholavaram",
      "Sirugavoor",
      "Sothuperumbedu",
      "Thirumanam",
      "Thirunindravur B",
      "Thiruninravur",
      "Thiruninravur-A",
      "Thiruninravur-B",
      "Thiruvotriyur",
      "Tondairpet",
      "Tondiarpet",
      "Vanagaram",
      "Vayalanallur",
      "Vayalanallur-A",
      "Veeraragavapuram",
      "Veeraraghavapuram",
      "Venkatapuram",
      "Vilangadupakkam",
      "Villivakkam",
      "Ward No. I of Paruthipattu"
    ]
  },
  {
    "Ventakesh / Karthik": [
      "Arasankalani",
      "Arasankazhani"
    ]
  },
  {
    "Karthik / Ventakesh": [
      "Gerugambakkam",
      "Kollacheri",
      "Kulappakkam",
      "Kuthambakkam",
      "Poonamallee",
      "Rendamkattalai",
      "Rendankattalai",
      "Sikkarayapuram",
      "Vellavedu",
      "Zamin Pallavaram",
      "Zamin Pallvaram"
    ]
  },
  {
    "Karthik / Jagan": [
      "Mambalam"
    ]
  },
  {
    "Jagan / karthik": [
      "Mylapore",
      "T. Nagar",
      "T.Nagar"
    ]
  },
  {
    "Ventakesh / Dinikaran": [
      "Part Kottivakkam",
      "Semmancheri",
      "Semmanchery"
    ]
  },
  {
    "Jagan /  Balachander": [
      "Sholavaram",
      "Sirugavoor"
    ]
  }
]