import pandas as pd
import requests
import time
from urllib.parse import urlparse

# Bing Web Search API Configuration
BING_API_KEY = '16c06aa8c818ffe24c264d66285e15925ae73da7c1d41318d73c195cdd3b3aa8'  # Replace with your actual API key
BING_SEARCH_ENDPOINT = 'Https:///search?engine=bing'

# Load the Excel file
file_path = 'C:/Users/kayla/Desktop/nonprofits_info.xlsx'
sheet_name = 'FL'

# Read the Excel sheet into a DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Columns are already correctly named - no changes needed
# Just ensure NaN values are handled properly (leave as NaN, we'll check with pd.isna)

def search_bing(query, count=5):
    """
    Search Bing and return results
    """
    headers = {'Ocp-Apim-Subscription-Key': BING_API_KEY}
    params = {
        'q': query,
        'count': count,
        'responseFilter': 'Webpages'
    }
    
    try:
        response = requests.get(BING_SEARCH_ENDPOINT, headers=headers, params=params)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error searching Bing: {e}")
        return None

def extract_homepage(org_name, state='Florida'):
    """
    Find organization homepage by excluding common directories and LinkedIn
    """
    # Search with organization name and state for better accuracy
    query = f'"{org_name}" {state}'
    results = search_bing(query, count=10)
    
    if not results or 'webPages' not in results:
        return ''
    
    # Domains to exclude when looking for homepage
    exclude_domains = [
        'linkedin.com', 'facebook.com', 'twitter.com', 'instagram.com',
        'guidestar.org', 'charitynavigator.org', 'candid.org',
        'wikipedia.org', 'irs.gov', 'sec.gov'
    ]
    
    for page in results['webPages']['value']:
        url = page['url']
        domain = urlparse(url).netloc.lower()
        
        # Skip excluded domains
        if any(excluded in domain for excluded in exclude_domains):
            continue
        
        # Return the first non-excluded result
        return url
    
    return ''

def extract_linkedin(org_name):
    """
    Find organization LinkedIn profile
    """
    query = f'"{org_name}" site:linkedin.com/company'
    results = search_bing(query, count=5)
    
    if not results or 'webPages' not in results:
        return ''
    
    for page in results['webPages']['value']:
        url = page['url']
        # Make sure it's a company page, not a personal profile or post
        if 'linkedin.com/company/' in url.lower():
            # Clean up the URL (remove query parameters)
            clean_url = url.split('?')[0]
            return clean_url
    
    return ''

def find_urls(org_name, state='Florida'):
    """
    Find both homepage and LinkedIn URL for an organization
    """
    print(f"Searching for: {org_name}")
    
    homepage = extract_homepage(org_name, state)
    time.sleep(0.5)  # Small delay between searches to be respectful
    
    linkedin = extract_linkedin(org_name)
    time.sleep(0.5)  # Small delay between searches
    
    return homepage, linkedin

# Process each organization
total_orgs = len(df)
for index, row in df.iterrows():
    org_name = row['Organization Name']
    
    # Check if URLs exist - pd.isna() will correctly detect NaN values
    homepage_val = df.at[index, 'Homepage URL']
    linkedin_val = df.at[index, 'LinkedIn URL']
    
    homepage_exists = pd.notna(homepage_val) and str(homepage_val).strip() != ''
    linkedin_exists = pd.notna(linkedin_val) and str(linkedin_val).strip() != ''
    
    # Skip if already has both URLs
    if homepage_exists and linkedin_exists:
        print(f"Skipping {org_name} (already has URLs)")
        continue
    
    print(f"\nProcessing {index + 1}/{total_orgs}: {org_name}")
    
    try:
        homepage, linkedin = find_urls(org_name)
        
        # Update with results (use None for empty, which Excel will show as blank)
        if homepage:
            df.at[index, 'Homepage URL'] = homepage
            print(f"  Homepage: {homepage}")
        else:
            df.at[index, 'Homepage URL'] = None
            print(f"  Homepage: Not found")
        
        if linkedin:
            df.at[index, 'LinkedIn URL'] = linkedin
            print(f"  LinkedIn: {linkedin}")
        else:
            df.at[index, 'LinkedIn URL'] = None
            print(f"  LinkedIn: Not found")
        
        # Save progress every 10 organizations
        if (index + 1) % 10 == 0:
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"\n--- Progress saved at {index + 1}/{total_orgs} ---\n")
            
    except Exception as e:
        print(f"  Error processing {org_name}: {e}")
        df.at[index, 'Homepage URL'] = None
        df.at[index, 'LinkedIn URL'] = None

# Final save
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name=sheet_name, index=False)
print("\n✓ All done! Results saved to Excel file.")