from pathlib import Path

# --- FILE PATHS ---
# Define base directory and paths for input/output files
BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_DIR = BASE_DIR / "input"
EXCEL_PATH = INPUT_DIR / "mock_databases" / "Power_Query_Database_mock.xlsx"

PDF_PATH = INPUT_DIR / "data_form_editable.pdf"

OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)



# --- CONFIGURATION ---
# Constant values that will be inserted into specific PDF fields
CONFIG = {
    "director": "John Smith",
    "exam_center_city": "Naples",
    "exam_center_country": "Italy",
    "institute_city": "Naples",
    "location": "Naples",
    "chunk_size":10
}

# --- MAPPING DEFINITIONS ---
# Maps PDF checkbox fields to their corresponding Excel columns
CHECKBOX_MAP = {
    "chk_gender": "Gender_chk",
    "chk_name": "Name_chk",
    "chk_surname": "Surname_chk",
    "chk_date_of_birth": "Date_of_birth_chk",
    "chk_place_of_birth": "Place_of_birth_chk",
    "chk_country_of_birth": "Country_of_birth_chk",
    "chk_email": "Email_chk",
}

# Maps PDF text fields to Excel columns.
# If the value is None, the field will be filled from CONFIG or current date.
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