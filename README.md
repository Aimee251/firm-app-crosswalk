# firm-app-crosswalk
1. Filter the companies that currently have mobile apps and companies which need further investigation
2. Parse the information that required for the targeted companies
- Search the Apple API by company name, get candidate apps, extract the developer ID, and use the developer ID to get all apps
under that developer, save everything to Excel
- Use the domain to match the apps' names and the company to solve inconsistent naming
- Use Claude API to verify the low-confidence apps
3. Based on the Apple app search results, scrape the Google Play apps to avoid using SerpApi, which isn't cost-efficient
   
# Packages installed to help solve the problem
- Pandas: data backbone
- openpyxl: read and write to Excel
- Requests: make HTTP calls to Apple's iTunes search
-  Antropic
-  python-dotenv: load my Claude key into the environment
