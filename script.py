import pandas as pd
import requests
import time
from openpyxl import load_workbook
import os

def get_virustotal_score(ip, api_key):
    """
    Check IP reputation on VirusTotal
    Returns: Number of malicious + suspicious detections
    """
    url = f"https://www.virustotal.com/api/v3/ip_addresses/{ip}"
    headers = {
        "x-apikey": api_key,
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        
        data = response.json()
        stats = data['data']['attributes']['last_analysis_stats']
        return stats['malicious'] + stats['suspicious']
        
    except requests.exceptions.RequestException as e:
        return f"Network Error: {str(e)}"
    except KeyError as e:
        return f"Data Error: {str(e)}"
    except Exception as e:
        return f"Error: {str(e)}"

def main():
    # Configuration
    API_KEY = "your API HERE"  # Replace with your API key
    EXCEL_FILE = "check.xlsx"       # Replace with your Excel file name
    SHEET_NAME = "check"                     # Change if your sheet has different name
    
    # Read Excel file
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        print(f"Loaded {len(df)} rows from {EXCEL_FILE}")
    except FileNotFoundError:
        print(f"Error: File {EXCEL_FILE} not found!")
        return
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return
    
    # Check if Column B exists, if not create it
    if 'Malicious Count' not in df.columns:
        df['Malicious Count'] = ''
    
    # Process IPs
    processed = 0
    total_ips = len(df)
    
    for index, row in df.iterrows():
        ip = row.iloc[0]  # Get IP from first column (Column A)
        
        # Skip empty cells and headers
        if pd.isna(ip) or ip == '' or str(ip).lower() in ['ip', 'ip address', 'ip_address']:
            continue
        
        # Check if already processed
        if pd.isna(row['Malicious Count']) or row['Malicious Count'] == '':
            print(f"Checking IP {processed + 1}/{total_ips}: {ip}")
            result = get_virustotal_score(str(ip).strip(), API_KEY)
            df.at[index, 'Malicious Count'] = result
            processed += 1
            
            # Save progress after each check (optional)
            df.to_excel(EXCEL_FILE, index=False, sheet_name=SHEET_NAME)
            
            # Respect VirusTotal rate limits (15 seconds for free API)
            if processed < total_ips:
                print(f"Waiting 15 seconds... ({processed}/{total_ips} completed)")
                time.sleep(15)
    
    # Final save
    df.to_excel(EXCEL_FILE, index=False, sheet_name=SHEET_NAME)
    print(f"Completed! Processed {processed} IP addresses.")
    print(f"Results saved to {EXCEL_FILE}")

if __name__ == "__main__":
    main()
