📋 Overview

This tool reads IP addresses from an Excel file, queries VirusTotal's API for each IP, and writes the results back to the Excel file. Perfect for security analysts, network administrators, and cybersecurity professionals who need to quickly assess the reputation of multiple IP addresses.

✨ Features

Automated IP Reputation Checking: Bulk check multiple IP addresses against VirusTotal

Excel Integration: Direct read/write from/to Excel files

Rate Limit Handling: Automatically respects VirusTotal's API rate limits (15 seconds between requests)

Error Handling: Robust error handling for network issues and invalid IPs

Progress Tracking: Real-time progress updates in the console

Resume Capability: Skips already processed IPs if the script is interrupted

📁 Files
script.py - Main Python script

check.xlsx - Example Excel file (Column A: IP addresses, Column B: Results)

🛠️ Prerequisites
Python 3.6+

VirusTotal API account (Sign up here)

Excel file with IP addresses in Column A

📦 Installation
Clone the repository:

bash
git clone [https://github.com/mohammadabidhossain/excel-ip-check-via-virustotal-API](https://github.com/mohammadabidhossain/excel-ip-check-via-virustotal-API-).git
cd excel-ip-check-via-virustotal-API
Install required packages:

bash
pip install pandas requests openpyxl
Get your VirusTotal API key:

Sign up at VirusTotal

Go to your account settings → API Key

Copy your API key

🚀 Usage
Prepare your Excel file:

Place IP addresses in Column A

Leave Column B empty for results

Save as check.xlsx in the same folder as the script

Configure the script:
Open script.py and replace the API key:

python
API_KEY = "your_actual_virustotal_api_key_here"
Run the script:

bash
python script.py
View results:

Open check.xlsx to see the results in Column B

Results show the number of engines that detected the IP as malicious/suspicious

📊 Excel Format
IP Address (Column A)	Malicious Count (Column B)
8.8.8.8	0
1.1.1.1	0
malicious.ip.com	15
Result Values:

0: Clean (no detections)

1+: Number of malicious/suspicious detections

Error: Could not check the IP

⚙️ Configuration
Edit these variables in script.py:

python
API_KEY = "your_api_key_here"        # Your VirusTotal API key
EXCEL_FILE = "check.xlsx"            # Your Excel filename
SHEET_NAME = "Sheet1"                # Excel sheet name (default: Sheet1)
REQUEST_DELAY = 15                   # Delay between requests in seconds
🔄 API Rate Limits
Free API: 4 requests per minute (15-second delay enforced)

Paid API: Remove or reduce the time.sleep(15) line for higher limits

🐛 Troubleshooting
Common Issues:

"Module not found" errors:

bash
pip install pandas requests openpyxl
Excel file permission errors:

Close Excel before running the script

API rate limit errors:

Ensure you're using a valid API key

The script automatically handles rate limiting

Network errors:

Check your internet connection

Verify the API key is correct

📝 Example Output
text
Found 10 IP addresses to check...
Checking: 8.8.8.8
Result: 0
Waiting 15 seconds...
Checking: 1.1.1.1
Result: 0
Waiting 15 seconds...
...
All IPs processed! Check your Excel file.
🤝 Contributing
Fork the repository

Create a feature branch: git checkout -b feature-name

Commit changes: git commit -am 'Add new feature'

Push to branch: git push origin feature-name

Submit a pull request

📄 License
This project is licensed under the MIT License - see the LICENSE file for details.

⚠️ Disclaimer
This tool is for educational and legitimate security purposes only. Ensure you have proper authorization to scan any IP addresses. The authors are not responsible for any misuse of this tool.

🙏 Acknowledgments
VirusTotal for providing the API

Python community for the excellent libraries used
