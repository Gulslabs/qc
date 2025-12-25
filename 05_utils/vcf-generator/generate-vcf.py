import pandas as pd

# Load your CSV or Excel file
df = pd.read_csv("contacts.csv")  # or pd.read_excel("contacts.xlsx")

def format_mobile(mobile):
    mobile = str(mobile).strip().replace(" ", "").replace("-", "")
    if not mobile.startswith("+91"):
        if mobile.startswith("91"):
            return f"+{mobile}"
        return f"+91{mobile}"
    return mobile

with open("contacts.vcf", "w", encoding="utf-8") as vcf_file:
    for index, row in df.iterrows():
        name = row["Name"]
        mobile = format_mobile(row["Mobile"])
        
        # Write vCard entry
        vcf_file.write("BEGIN:VCARD\n")
        vcf_file.write("VERSION:3.0\n")
        vcf_file.write(f"N:{name};;;\n")
        vcf_file.write(f"FN:{name}\n")
        vcf_file.write(f"TEL;TYPE=CELL:{mobile}\n")
        vcf_file.write("END:VCARD\n\n")
