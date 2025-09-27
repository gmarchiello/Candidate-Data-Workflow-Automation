import os, re, sys, shutil
import pandas as pd
from datetime import datetime
from pytz import timezone
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfName, PdfObject
from copy import deepcopy

# --- File paths ---
EXCEL_PATH = "/Users/g/Documents/Data_change_portfolio/data/personal_data_change.xlsx" # PUT FILE PATHS BEFORE CONFIG
PDF_TEMPLATE = "/Users/g/Documents/Data_change_portfolio/data/data_form_editable.pdf"

# --- Config dictionary for fixed values ---
CONFIG = {
    "director": "John Smith",
    "exam_center_city": "Naples",
    "exam_center_country": "Italy",
    "institute_city": "Naples",
    "location": "Naples",
}

# --- Checkbox mapping: PDF checkbox -> Excel column ---
CHECKBOX_MAP = {
    "chk_gender": "Gender_chk",
    "chk_name": "Name_chk",
    "chk_surname": "Surname_chk",
    "chk_date_of_birth": "Date_of_Birth_chk",
    "chk_place_of_birth": "Place_of_birth_chk",
    "chk_country_of_birth": "Country_of_birth_chk",
    "chk_email": "Email_chk",
}

# --- Text field mapping: PDF field -> Excel column or None for fixed values ---
TEXT_MAP = {
    "txt_director": None,
    "txt_exam_center_city": None,
    "txt_exam_center_country": None,
    "txt_institute_city": None,
    "txt_client_code": "Client_code",
    "txt_exam_code": "Exam_code",
    "txt_gender": "Gender",
    "txt_name": "Name",
    "txt_surname": "Surname",
    "txt_country_of_birth": "Country_of_birth",
    "txt_date_of_birth": "Date_of_birth",
    "txt_place_of_birth": "Place_of_birth",
    "txt_email": "Email",
    "txt_location": None,
    "txt_today_date": None,
}

# --- Load Excel and PDF safely ---
try:

  #check if excel path exist
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel file not found at: {EXCEL_PATH}")
    if not os.path.exists(PDF_TEMPLATE):
        raise FileNotFoundError(f"PDF template not found at: {PDF_TEMPLATE}")

  #check if excel columns are valid
    df = pd.read_excel(EXCEL_PATH)
    df.columns = df.columns.str.strip()
    if df.columns.isnull().any() or (df.columns == "").any():
        raise ValueError("Excel contains empty or invalid column headers.") # Remove here X ERROR

    print(f"✅ Excel loaded successfully: {EXCEL_PATH}")
    print(f"✅ Column headers normalized and valid: {list(df.columns)}")

# ADD A CHECK TO SEE IF EXCEL COLUMNS ARE THE ONES IN THE MAP


  #check if pdf has form fields

    pdf_template = PdfReader(PDF_TEMPLATE)
    if not pdf_template.Root.AcroForm or not getattr(pdf_template.Root.AcroForm, "Fields", None):
        raise ValueError("PDF template does not contain form fields.")
    print(f"✅ PDF template loaded successfully: {PDF_TEMPLATE}")

    # Check PDF contains all required fields
    pdf_fields = set()
    for field in pdf_template.Root.AcroForm.Fields:
        if field.T:
            pdf_fields.add(field.T.to_unicode().strip())  # check why to_unicode?

    missing_text = set(TEXT_MAP.keys()) - pdf_fields
    missing_checkboxes = set(CHECKBOX_MAP.keys()) - pdf_fields

    if missing_text or missing_checkboxes:
        raise ValueError(
            f"PDF template is missing required fields.\n"
            f"Missing text fields: {missing_text}\n"
            f"Missing checkboxes: {missing_checkboxes}"
        )
    else:
        print("✅ PDF contains all required fields.")

#--- OUT HERE MORE EXCEPT CASES FOR SPECIFIC EXCETIONS: ValueError, FileNotFoundError
except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
    sys.exit(1)

except ValueError as e:
    print(f"❌ ERROR: {e}")
    sys.exit(1)

except Exception as e:
    print(f"❌ ERROR: {e}")
    sys.exit(1)

# --- OUTPUT FOLDER ---
italy_tz = timezone("Europe/Rome")
timestamp = datetime.now(italy_tz).strftime("%Y%m%d_%H%M")
BASE_OUTPUT_DIR = "/Users/g/Documents/Data_change_portfolio/output"
OUTPUT_FOLDER = os.path.join(BASE_OUTPUT_DIR, f"fulfilled_forms_{timestamp}")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
print(f"✅ Output folder created: {OUTPUT_FOLDER}")

# --- Get changed fields from Excel row --- check which column got the "ON", retunr a friendly name for the messages etc
def get_changed_fields(row):
    changed = []
    for pdf_checkbox, excel_col in CHECKBOX_MAP.items(): #PUT HERE LIKE "pdf_checkbox, excel_col" BECAUSE YOU CHANGE IT IN MAPPING
        if str(row.get(excel_col, "")).strip().upper() == "ON":
            friendly_name = excel_col.replace("_chk", "")
            changed.append(friendly_name)
    return changed

# --- Safe getter --- For file name and pdf field - IN CASE FIELD IS EMPY, RETUNR "" in PDF FIELD OR UNKNOWN in the name of the file 
def safe_get(value, for_pdf_field=True, placeholder="UNKNOWN"):
    if pd.isna(value) or str(value).strip() in ["", "nan", "NaT"]:
        return "" if for_pdf_field else placeholder
    return str(value).strip()

# --- Fill PDF function ---
def fill_pdf(input_pdf_path, output_pdf_path, text_values, checkboxes_to_check): # Rename "input_pdf_path" to "pdf" (since it will be an input value)
   pdf = PdfReader(PDF_TEMPLATE)  # MOVE THIS TO "Main loop"
    #maybe better put deep copy
   pdf.Root.AcroForm.update({PdfName("NeedAppearances"): PdfObject("true")}) # special trick in PDF forms that ensures the filled-in values actually appear correctly when you open the PDF. Let me explain carefully.

   for page in pdf.pages:
        annotations = page.Annots
        if not annotations:
            continue

        for annot in annotations:
            if annot.Subtype == PdfName.Widget and annot.T: #annot.T truely?
                key = annot.T.to_unicode().strip()
                #Text fields
                if key in text_values:
                    annot.V = text_values[key]
                    annot.AP = None #what is annot.ap?
                #Checkboxes
                elif key in checkboxes_to_check: # Change "if" to "elif"
                    annot.V = PdfName("Yes")
                    annot.AS = PdfName("Yes")

        PdfWriter().write(output_pdf_path, pdf)

# --- Main loop ---
email_data = []
# STEP 1 MAINLOOP - create list of MISSING FIELDS, CHANGE FIELDS; MISSING CHECKBOXES
#Let's start the loop row by row
for idx, row in df.iterrows():
    #for each column, give me the fields 
    missing_fields = [
        excel_col.upper().replace(" ", "_")         # Normalize field names to match expected format (uppercase and underscores) for consistency with PDF field naming
        for excel_col in TEXT_MAP.values() #get the values of the text map, so the excel columns
        if excel_col is not None and (pd.isna(row.get(excel_col, None)) or str(row.get(excel_col)).strip() == "") # check if the excel_col is not null and if the value in the row is null or empty
      ]
    changed_fields = get_changed_fields(row) # returns a list of friendly field names for checked checkboxes
    missing_checkbox = not bool(changed_fields) #what is not bool?



    # --- Filename suffix --- create SUFFIX to use in  the file name, including MISSING FIELDS, CHANGE FIELDS; MISSING CHECKBOXES- 
    suffix_list = []
    if changed_fields:
        suffix_list.append("_".join(changed_fields).replace(" ", "_")) #from the changed fields we remove _ and add a space and we append it in the suffix list
    if missing_fields:
        suffix_list.append("MISSING_" + "_".join(missing_fields)) #if there is a missing field, we append it in the suffix list
    if missing_checkbox:
        suffix_list.append("MISSING_CHECKBOX") # if we have some missing checkbox we append it in the suffix list, but we don't remove the underscore???

   # EXTRACT NAME, SURNAME, CREATE A SAFENAME FOR THE FILE (JOINING THE SUFFIXES) 
    name = safe_get(row.get("Name"), for_pdf_field=False) #we take the name from the column
    surname = safe_get(row.get("Surname"), for_pdf_field=False) #we take the surname from the column
    safe_name = re.sub(r"[^a-zA-Z0-9]", "_", f"{surname}_{name}_change_request_" + "_".join(suffix_list)) #we create the file name, check the regular expressions
    output_pdf_path = os.path.join(OUTPUT_FOLDER, f"{safe_name}.pdf") # check this part, how this os.path works and the join

    # --- Build PDF data dynamically --- creates a dictonary of values that will be filled into the pdf for each row of the excel file
    text_values = {} #initialize a dictionary
    #assign the text values
    for pdf_field, excel_col in TEXT_MAP.items(): #itereate through text map
        if excel_col:  # if there is a value in the excel textmap, and is not a Null, is it a trully?
            value = row.get(excel_col)  #check how row.get works. row.get(key, default=None) is a method of Pandas Series (similar to a Python dictionary) that Tries to return the value corresponding to key (column name). If the key does not exist, it returns default (which defaults to None).
            if pdf_field == "txt_date_of_birth":
                birth_date = pd.to_datetime(value, errors="coerce")
                text_values[pdf_field] = birth_date.strftime("%d/%m/%Y") if pd.notnull(birth_date) else "" #Python’s ternary (inline if-else) expressio
            else:
                text_values[pdf_field] = safe_get(value) #IS IT A MISTAKE?
        else:  # fixed value from CONFIG or current date
            if pdf_field in ["txt_director", "txt_exam_center_city", "txt_exam_center_country",
                             "txt_institute_city", "txt_location"]:
                text_values[pdf_field] = CONFIG[pdf_field.replace("txt_", "")]
            elif pdf_field == "txt_today_date":
                text_values[pdf_field] = datetime.now(italy_tz).strftime("%d %B %Y")

    checkboxes_to_check = []
    for pdf_field, excel_col in CHECKBOX_MAP.items():
        if str(row.get(excel_col, "")).strip().upper() == "ON": #Default value is used in case the column do not exists
            checkboxes_to_check.append(pdf_field)
    

    # --- Fill PDF and add to ZIP ---
    try:
        fill_pdf(PDF_TEMPLATE, output_pdf_path, text_values, checkboxes_to_check) 
    except Exception as e:   #try to understand how this work. : catch errors that happen while filling a specific PDF. Examples of row-specific errors: Invalid date in this row Special character in a name that crashes PDF write Permission issues writing/removing this particular fil 
        print(f"❌ Error processing row {idx}: {e}")
        continue                 #The difference: it does not stop the whole script; it logs the error and moves to the next row.


    # --- Logging ---
    if missing_fields and missing_checkbox:
        print(f"⚠️ Row {idx} missing text value: {', '.join(missing_fields)}")
        print(f"❗ Row {idx} has no checkboxes selected")
        print(f"🟡 Row {idx} processed: {os.path.basename(output_pdf_path)}")
    elif missing_fields:
        print(f"⚠️ Row {idx} missing text value: {', '.join(missing_fields)}")
        print(f"🟡 Row {idx} processed: {os.path.basename(output_pdf_path)}")
    elif missing_checkbox:
        print(f"❗ Row {idx} has no checkboxes selected")
        print(f"🔴 Row {idx} processed: {os.path.basename(output_pdf_path)}")
    else:
        print(f"✔️ Row {idx} processed: {os.path.basename(output_pdf_path)}")

    # --- Prepare email data ---
    changed_text = ", ".join(changed_fields) if changed_fields else "No changes"
    subj_name = f"{surname} {name}".strip()
    body_line = f"{surname} {name} ({changed_text})"
    if missing_fields:
        body_line += " ⚠️ MISSING " + ", ".join([mf.replace("_", " ") for mf in missing_fields])
    if missing_checkbox:
        body_line += " ❗ MISSING CHECKBOX"
    email_data.append((surname, name, subj_name, body_line))


# --- Generate email messages ---
email_data_with_changes = [
    entry for entry in email_data
    if "No changes" not in entry[3] or "❗ MISSING CHECKBOX" in entry[3] or "⚠️ MISSING" in entry[3]
]
email_data_with_changes.sort()
chunk_size = 10
chunks = [email_data_with_changes[i:i + chunk_size] for i in range(0, len(email_data_with_changes), chunk_size)]

for i, chunk in enumerate(chunks):
    subject_names = [entry[2] for entry in chunk]
    subject = f"Change Request: {', '.join(subject_names)}"
    body_lines = [f"- {entry[3]}" for entry in chunk]
    body = "Good morning,\nI kindly ask you to update the data of the following candidates:\n\n" + "\n".join(body_lines)
    print(f"\n--- Message {i + 1} ---")
    print(subject)
    print("\n" + body)
