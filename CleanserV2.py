import pandas as pd
import re
import os

def format_job_title(title):
    """Format job titles based on the given rules."""
    if isinstance(title, str):

        if "Director" in title:
            # Shorten "Chief Information Security Officer" to "CISO" if present
            if "Chief Information Security Officer" in title or "(CISO)" in title or "CISO" in title:
                title = title.replace("Chief Information Security Officer", "")
                title = title.replace("(CISO)", "")
                title = title.replace("CISO", "")
                title = title.strip()
                title = f"{title} CISO".strip()

            # Shorten "Head of IT Security" and "Head of IT" if they fit the conditions
            if "Head" in title and ("IT" in title or "Information Technology" in title):
                if "cyber" in title.lower() or "security" in title.lower():
                    title = re.sub(r"Head.*?IT.*?(Cyber|Security).*", "Head of IT Security", title, flags=re.IGNORECASE)
                else:
                    title = re.sub(r"Head.*?IT.*", "Head of IT", title, flags=re.IGNORECASE)
            return title

        # If "Director" is not present, perform the same changes without skipping
        if "CISO" in title or "Chief Information Security Officer" in title:
            # Nested check for "CIO"
            if "CIO" in title:
                return "CIO and CISO"
            return "CISO"

        # Shorten "Head of IT Security" and "Head of IT"
        if "Head" in title and ("IT" in title or "Information Technology" in title):
            if "cyber" in title.lower() or "security" in title.lower():
                return "Head of IT Security"
            else:
                return "Head of IT"

        # Capitalize correctly
        exceptions = {"of", "and", "in", "on", "at", "for", "to"}
        title = ' '.join(word.capitalize() if word.lower() not in exceptions else word.lower()
                         for word in title.split())

        # Ensure IT is fully capitalized when it is a standalone word
        title = ' '.join(word if word.lower() != "it" else "IT" for word in title.split())

    return title

def format_company_name(name):
    """Format company names based on the given rules."""
    if isinstance(name, str):
       # Remove everything in parentheses along with the content inside, including nested parentheses
        stack = []
        cleaned_name = ""
        for char in name:
            if char == '(':
                stack.append('(')  # Track open parentheses
            elif char == ')' and stack:
                stack.pop()  # Close the most recent open parentheses
            elif not stack:
                cleaned_name += char  # Only add characters outside parentheses
        name = cleaned_name.strip()

        # Remove unnecessary suffixes
        for suffix in ["Inc", "Corporation", "Corp", "LLC", "Ltd", "Limited"]:
            name = name.split(suffix)[0]

        # Remove punctuation and trim
        name = name.replace(",", "").replace("-", "").strip()

        # Capitalize correctly (ignore all-caps words as they might be abbreviations)
        name = ' '.join(word if word.isupper() else word.capitalize() for word in name.split())
    return name

def process_excel(input_file):
    """Process the Excel file to edit Job Title and Company Name columns."""
    # Load the Excel file
    df = pd.read_excel(input_file)

    # Edit the Job Title column
    if "Job Title" in df.columns:
        df["Job Title"] = df["Job Title"].apply(format_job_title)

    # Edit the Company Name column
    if "Company Name" in df.columns:
        df["Company Name"] = df["Company Name"].apply(format_company_name)

    # Save the normalized data to 'email [filename].xlsx'
    base_filename = os.path.splitext(os.path.basename(input_file))[0]
    email_output_file = f"email_{base_filename}.xlsx"
    df.to_excel(email_output_file, index=False, engine='openpyxl')
    print(f"Normalized data saved to: {email_output_file}")

    # Keep only relevant columns for the second file
    columns_to_keep = ["Linkedin Url", "First Name", "Job Title", "Company Name"]
    df_filtered = df[columns_to_keep]

    # Save the filtered data to 'LK [filename].csv' with UTF-8 encoding
    lk_output_file = f"LK_{base_filename}.csv"
    df_filtered.to_csv(lk_output_file, index=False, encoding='utf-8-sig')
    print(f"Filtered data saved to: {lk_output_file}")

# File paths
input_file = 'Apollo_Data_20250108114847.xlsx'

# Process the file
process_excel(input_file)
