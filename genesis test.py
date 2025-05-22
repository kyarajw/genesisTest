#import datetime module, pandas library
import datetime
import pandas as pd

#dictionary to track base_id appearance count
id_counter = {}

#calculates FY from date
def get_fy_from_date(date):
    if pd.isna(date):
        date = datetime.datetime.now()
    elif isinstance(date, str):
        try:
            date = pd.to_datetime(date, dayfirst=True)
        except:
            date = datetime.datetime.now()
    
    # Get last 2 digits of year, adding 1 if after March
    year = date.year + (1 if date.month >= 4 else 0)
    return str(year)[2:]

#function to generate unique ids for each case
def make_id(org, referral_date):
    # Get FY from referral date
    year = get_fy_from_date(referral_date)
    
    #name handling to ensure valid input is passed  
    if pd.isna(org):
        org = "unknown"
    else:
        org = str(org)
    
    #convert ref org name to lowercase initials string
    org = org.lower()
    org = org.replace('(', '').replace(')', '')
    org = ''.join(word[0] for word in org.split() if word)
    
    #creates base id (org_FY)
    #calculates num component based on if/how many times base_id has previously appeared
    base_id = f"{org}_FY{year}"
    if base_id in id_counter:
        id_counter[base_id] = id_counter[base_id] + 1
    else:
        id_counter[base_id] = 0
    
    #combines strings to create unique id
    new_id = f"{base_id}_{id_counter[base_id]:03d}"
    return new_id

#read excel file from row 7
excel_file = "Exemplar Report.xlsx"
df = pd.read_excel(excel_file, header=6) 

#iterates through every row, calling make_id with org name and ref. date args
#stores output in new id column in df
df['ID'] = df.apply(lambda row: make_id(row.iloc[2], row.iloc[3]), axis=1)

#save to output file
df.to_excel("Output.xlsx", index=False)
print("Output file saved as Output.xlsx")
        
