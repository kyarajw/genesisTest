import pandas as pd
from datetime import datetime

# Load Excel file, treating row 7 as the header
excel_file = "Exemplar Report Youth Genesis Project Magma.xlsx"
df = pd.read_excel(excel_file, skiprows=6)

# Normalize column names to lowercase and strip spaces
df.columns = [col.strip().lower() for col in df.columns]

# Define possible DOB column names
dob_cols = ['dob', 'date of birth']

# Check for presence of 'age' and possible DOB columns
has_age = 'age' in df.columns
has_dob = any(col in df.columns for col in dob_cols)

# Store today's year for age calculation
today_year = datetime.today().year

# Apply logic
# If the data doesn't have "age" nor "DOB" column
if not has_age and not has_dob:
    df['age'] = pd.NA
    print("Neither 'Age' nor 'DOB' exists. Created 'age' column with NA.")

# If the data only has "DOB" or "date of birth" column
elif not has_age and has_dob:
    dob_col = [col for col in dob_cols if col in df.columns][0]
    df[dob_col] = pd.to_datetime(df[dob_col], errors='coerce')
    df['age'] = today_year - df[dob_col].dt.year
    df.drop(columns=[dob_col], inplace=True)
    print(f"'Age' calculated from '{dob_col}' and '{dob_col}' column dropped.")

# If the data has "age" and "DOB" column
elif has_age and has_dob:
    dob_col = [col for col in dob_cols if col in df.columns][0]
    df.drop(columns=[dob_col], inplace=True)
    print(f"'Age' exists. Dropped redundant '{dob_col}' column.")

# Identify empty 'age' rows (reflecting original Excel row numbers)
if 'age' in df.columns:
    empty_rows = df[df['age'].isna()].index + 7  # Excel rows start at row 7
    if not empty_rows.empty:
        print("Empty 'age' values found at Excel rows:")
        print(empty_rows.tolist())


# Capitalize column names before saving
df.columns = df.columns.str.title()

# Save to Excel
df.to_excel("Output.xlsx", index=False)
print("Output file saved as 'Output.xlsx' with capitalized column names.")
