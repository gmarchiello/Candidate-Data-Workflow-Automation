import os, re, sys
import pandas as pd
from datetime import datetime
from pytz import timezone
from pdfrw import PdfReader
from config import CONFIG, CHECKBOX_MAP, TEXT_MAP, EXCEL_PATH, PDF_PATH, OUTPUT_DIR
from utils import safe_get, make_output_folder, clean_filename, get_checked_fields
from pdf_filler import fill_pdf


# --- TIMEZONE SETUP ---
italy_tz = timezone("Europe/Rome")

# --- LOAD AND VALIDATE FILES ---
try:
    # Check if Excel and PDF files exist, raise error if missing
    missing_files = []
    if not EXCEL_PATH.exists():
        missing_files.append(str(EXCEL_PATH))
    if not PDF_PATH.exists():
        missing_files.append(str(PDF_PATH))
    if missing_files:
        file_word = "file is" if len(missing_files) == 1 else "files are"
        raise FileNotFoundError(
        f"The following required {file_word} missing:\n - " + 
        "\n - ".join(str(f) for f in missing_files)
        )
    print("✅ All required files are present.")

    # Load Excel data and normalize column headers
    df = pd.read_excel(str(EXCEL_PATH), sheet_name = "PYTHON")
    df.columns = df.columns.str.strip()
    if df.columns.isnull().any() or (df.columns == "").any():
        raise ValueError("Excel contains empty or invalid column headers.")

    print(f"✅ Excel loaded successfully: {EXCEL_PATH}")
    print(f"✅ Column headers normalized and valid: {list(df.columns)}")

    # Load PDF template and validate existence of form fields
    pdf_template = PdfReader(str(PDF_PATH))
    if not pdf_template.Root.AcroForm or not getattr(pdf_template.Root.AcroForm, "Fields", None):
        raise ValueError("PDF template does not contain form fields.")
    print(f"✅ PDF template loaded successfully: {PDF_PATH}")

    # Collect all available PDF field names for validation
    pdf_fields = set()
    for field in pdf_template.Root.AcroForm.Fields:
        if field.T:
            pdf_fields.add(field.T.to_unicode().strip())

    # Ensure all required mapped fields exist in the PDF template
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

    # --- Validate Excel columns ---
    excel_cols = set(df.columns)

    # Columns required from TEXT_MAP and CHECKBOX_MAP (exclude None values)
    required_text_cols = set(c for c in TEXT_MAP.values() if c is not None)
    required_checkbox_cols = set(CHECKBOX_MAP.values())

    missing_text_cols = required_text_cols - excel_cols
    missing_checkbox_cols = required_checkbox_cols - excel_cols

    if missing_text_cols or missing_checkbox_cols:
        raise ValueError(
            f"Excel is missing required columns.\n"
            f"Missing text columns: {missing_text_cols}\n"
            f"Missing checkbox columns: {missing_checkbox_cols}"
        )
    else:
        print("✅ Excel contains all required columns.")


# Handle errors clearly and stop execution if critical
except FileNotFoundError as e:
    print(f"❌ File not found. Please check your file paths.\nError details: {e}")
    sys.exit(1)
except ValueError as e:
    print(f"❌ Configuration error. Please check your files.\nError details: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ An unexpected error occurred.\nError details: {e}")
    sys.exit(1)


# --- OUTPUT FOLDER ---
# Create timestamped output directory for generated PDFs"
OUTPUT_FOLDER = make_output_folder(str(OUTPUT_DIR))  # should return Path


# --- PROCESS EXCEL ROWS AND GENERATE FILLED PDFS ---
email_data = []

for idx, row in df.iterrows():


    
    # Identify selected checkboxes
    checked_fields = get_checked_fields(row, CHECKBOX_MAP)
    missing_checkbox = not bool(checked_fields)

        # Identify missing text fields
    missing_fields = [
        excel_col for excel_col in TEXT_MAP.values() # only consider mapped columns
        if excel_col is not None and not safe_get(row.get(excel_col)) # check for missing
    ]

    # Build descriptive filename suffix
    suffix_list = []
    if checked_fields:
        suffix_list.append("_".join(checked_fields).replace(" ", "_"))
    if missing_fields:
        suffix_list.append("MISSING_" + "_".join(missing_fields))
    if missing_checkbox:
        suffix_list.append("MISSING_CHECKBOX")

    # Construct safe output filename
    name = safe_get(row.get("Name"), for_pdf_field=False)
    surname = safe_get(row.get("Surname"), for_pdf_field=False)
    safe_name = clean_filename(name, surname, suffix_list)
    output_pdf_path = OUTPUT_FOLDER / f"{safe_name}.pdf"

    # Prepare text values for PDF fields
    text_values = {}
    for pdf_field, excel_col in TEXT_MAP.items():
        if excel_col:
            value = row.get(excel_col)
            if pdf_field == "txt_date_of_birth":
                # Format date of birth in dd/mm/yyyy
                birth_date = pd.to_datetime(value, errors="coerce")
                text_values[pdf_field] = birth_date.strftime("%d/%m/%Y") if pd.notnull(birth_date) else ""
            else:
                text_values[pdf_field] = safe_get(value)
        else:
            # Use fixed config values or today's date where applicable
            if pdf_field in ["txt_director", "txt_exam_center_city", "txt_exam_center_country",
                             "txt_institute_city", "txt_location"]:
                text_values[pdf_field] = CONFIG[pdf_field.replace("txt_", "")]
            elif pdf_field == "txt_today_date":
                text_values[pdf_field] = datetime.now(italy_tz).strftime("%d %B %Y")

    # Identify checkboxes to mark
    checkboxes_to_check = []
    for pdf_field, excel_col in CHECKBOX_MAP.items():
        if str(row.get(excel_col, "")).strip().upper() == "ON":
            checkboxes_to_check.append(pdf_field)

    # Fill PDF with data and handle row-specific errors gracefully
    try:
        fill_pdf(str(PDF_PATH), str(output_pdf_path), text_values, checkboxes_to_check)
    except Exception as e:
        print(f"❌ Error processing row {idx}: {e}")
        continue

    # Detailed row processing logs
    if missing_fields and missing_checkbox:
        print(f"⚠️ Row {idx} missing text value: {', '.join(missing_fields)}")
        print(f"❗ Row {idx} has no checkboxes selected")
        print(f"🟡 Row {idx} processed: {output_pdf_path.name}")
    elif missing_fields:
        print(f"⚠️ Row {idx} missing text value: {', '.join(missing_fields)}")
        print(f"🟡 Row {idx} processed: {os.path.basename(output_pdf_path)}")
    elif missing_checkbox:
        print(f"❗ Row {idx} has no checkboxes selected")
        print(f"🔴 Row {idx} processed: {os.path.basename(output_pdf_path)}")
    else:
        print(f"✔️ Row {idx} processed: {os.path.basename(output_pdf_path)}")

    # Prepare email-friendly summary for this row
    changed_text = ", ".join(checked_fields) if checked_fields else "No changes"
    body_line = f"{surname} {name} ({changed_text})"
    if missing_fields:
        body_line += " ⚠️ MISSING " + ", ".join([mf.replace("_", " ") for mf in missing_fields])
    if missing_checkbox:
        body_line += " ❗ MISSING CHECKBOX"
    email_data.append((surname, name, body_line))


# --- GENERATE EMAIL MESSAGES ---
# Only include rows with actual changes or missing data, excluding completely empty rows
email_data_with_changes = [
    entry for entry in email_data
    if "No changes" not in entry[2] or "❗ MISSING CHECKBOX" in entry[2] or "⚠️ MISSING" in entry[2]
]

# Sort and chunk results for grouping into manageable email messages
email_data_with_changes.sort()
chunk_size = CONFIG["chunk_size"]
chunks = [email_data_with_changes[i:i + chunk_size]
          for i in range(0, len(email_data_with_changes), chunk_size)]

# Build email subjects and bodies
for i, chunk in enumerate(chunks):
    subject_surname_name = [f"{entry[1]} {entry[0]}" for entry in chunk]
    email_subject = f"Change Request: {', '.join(subject_surname_name)}"
    email_body_lines = [f"- {entry[2]}" for entry in chunk]
    email_body = "Good morning,\nI kindly ask you to update the data of the following candidates:\n\n" + "\n".join(email_body_lines)
    print(f"\n--- Message {i + 1} ---")
    print(email_subject)
    print("\n" + email_body)
    print("\n-------------------\n")