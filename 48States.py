import pandas as pd
import requests
import time

print("Downloading IRS nonprofit data for all 48 continental states...")
print("This will take a few minutes...\n")

# Function to convert to sentence case (first letter capital, rest lowercase)
def to_sentence_case(text):
    if pd.isna(text) or text == '':
        return text
    text = str(text).strip()
    # Convert to title case (capitalizes first letter of each word)
    return text.title()

# All 50 states except Hawaii (HI) and Alaska (AK)
states = {
    'AL': 'Alabama', 'AZ': 'Arizona', 'AR': 'Arkansas', 'CA': 'California',
    'CO': 'Colorado', 'CT': 'Connecticut', 'DE': 'Delaware', 'FL': 'Florida',
    'GA': 'Georgia', 'ID': 'Idaho', 'IL': 'Illinois', 'IN': 'Indiana',
    'IA': 'Iowa', 'KS': 'Kansas', 'KY': 'Kentucky', 'LA': 'Louisiana',
    'ME': 'Maine', 'MD': 'Maryland', 'MA': 'Massachusetts', 'MI': 'Michigan',
    'MN': 'Minnesota', 'MS': 'Mississippi', 'MO': 'Missouri', 'MT': 'Montana',
    'NE': 'Nebraska', 'NV': 'Nevada', 'NH': 'New Hampshire', 'NJ': 'New Jersey',
    'NM': 'New Mexico', 'NY': 'New York', 'NC': 'North Carolina', 'ND': 'North Dakota',
    'OH': 'Ohio', 'OK': 'Oklahoma', 'OR': 'Oregon', 'PA': 'Pennsylvania',
    'RI': 'Rhode Island', 'SC': 'South Carolina', 'SD': 'South Dakota', 'TN': 'Tennessee',
    'TX': 'Texas', 'UT': 'Utah', 'VT': 'Vermont', 'VA': 'Virginia',
    'WA': 'Washington', 'WV': 'West Virginia', 'WI': 'Wisconsin', 'WY': 'Wyoming',
    'DC': 'District of Columbia'
}

# Create Excel writer
output_file = 'nonprofits_5m-15m_all_states.xlsx'
writer = pd.ExcelWriter(output_file, engine='openpyxl')

total_orgs = 0
state_counts = {}

for state_code, state_name in states.items():
    print(f"Processing {state_name} ({state_code})...")
    
    try:
        # Download state file from IRS
        url = f"https://www.irs.gov/pub/irs-soi/eo_{state_code.lower()}.csv"
        response = requests.get(url, timeout=30)
        
        if response.status_code == 200:
            # Save temporarily
            temp_file = f'temp_{state_code}.csv'
            with open(temp_file, 'wb') as f:
                f.write(response.content)
            
            # Read the CSV
            df = pd.read_csv(temp_file, encoding='latin1', low_memory=False)
            
            # Filter for revenue between $5M and $15M
            df_filtered = df[
                (df['INCOME_AMT'] >= 5000000) & 
                (df['INCOME_AMT'] <= 15000000)
            ].copy()
            
            if len(df_filtered) > 0:
                # Select only the columns we want
                output_df = df_filtered[['NAME', 'CITY', 'STATE', 'INCOME_AMT']].copy()
                output_df.columns = ['Organization Name', 'City', 'State', 'Annual Revenue']
                
                # Convert Organization Name and City to sentence case (title case)
                output_df['Organization Name'] = output_df['Organization Name'].apply(to_sentence_case)
                output_df['City'] = output_df['City'].apply(to_sentence_case)
                
                # Format revenue with dollar signs and commas
                output_df['Annual Revenue'] = output_df['Annual Revenue'].apply(
                    lambda x: f"${int(x):,}" if pd.notna(x) else ""
                )
                
                # Sort by revenue (highest to lowest)
                output_df = output_df.sort_values('Annual Revenue', ascending=False)
                
                # Write to Excel sheet
                # Excel sheet names have a 31 character limit
                sheet_name = state_code if len(state_code) <= 31 else state_code[:31]
                output_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                state_counts[state_name] = len(output_df)
                total_orgs += len(output_df)
                
                print(f"  ✓ {state_name}: {len(output_df):,} organizations")
            else:
                print(f"  ⚠ {state_name}: No organizations found in this revenue range")
                state_counts[state_name] = 0
        else:
            print(f"  ✗ {state_name}: Could not download data (status {response.status_code})")
            state_counts[state_name] = 0
        
        # Small delay to be respectful to IRS servers
        time.sleep(0.5)
        
    except Exception as e:
        print(f"  ✗ {state_name}: Error - {e}")
        state_counts[state_name] = 0

# Save the Excel file
writer.close()

print("\n" + "="*60)
print("COMPLETE!")
print("="*60)
print(f"\n✓ Total organizations across all states: {total_orgs:,}")
print(f"✓ Saved to: {output_file}")
print(f"✓ Total sheets: {sum(1 for v in state_counts.values() if v > 0)}")

print("\nTop 10 states by number of organizations:")
sorted_states = sorted(state_counts.items(), key=lambda x: x[1], reverse=True)
for i, (state, count) in enumerate(sorted_states[:10], 1):
    print(f"  {i}. {state}: {count:,} organizations")

input("\nPress Enter to close...")