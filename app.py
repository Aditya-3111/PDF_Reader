from flask import Flask, render_template, request, jsonify, send_file
import os, fitz, zipfile, io, re, mysql.connector, datetime
from openpyxl import Workbook
from dotenv import load_dotenv
import os

load_dotenv()  # load variables from .env file


app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY")

# ===== Database Configuration =====
DB_CONFIG = {
   "host": os.getenv("DB_HOST"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "database": os.getenv("DB_NAME")
}

# ====== Auto Database Setup ======
def init_db_auto():
    conn = mysql.connector.connect(
        host=DB_CONFIG["host"],
        user=DB_CONFIG["user"],
        password=DB_CONFIG["password"],
        auth_plugin='mysql_native_password'
    )
    cur = conn.cursor()
    cur.execute(f"CREATE DATABASE IF NOT EXISTS {DB_CONFIG['database']}")
    conn.commit()
    cur.close()
    conn.close()
    print("Database ready.")

init_db_auto()

# ===== PDF Extraction Logic =====
IGNORE_NAME_LINES = [
    "user information", "personal information", "form details",
    "applicant details", "registration form", "profile information"
]
NAME_PREFIXES = ["mr", "ms", "mrs", "dr", "prof"]

def extract_info_smart(text):
    data = {
        "filename": "Unknown", "name": "", "email": "", "phone": "",
        "company": "", "state": "", "district": "", "pin": ""
    }
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    email_match = re.findall(r"[\w\.-]+@[\w\.-]+\.\w+", text)
    if email_match:
        data["email"] = email_match[0]

    phone_match = re.findall(r"(\+?\d[\d\s\-]{7,}\d)", text)
    if phone_match:
        data["phone"] = phone_match[0].replace(" ", "")

    pin_match = re.findall(r"\b\d{6}\b", text)
    if pin_match:
        data["pin"] = pin_match[0]

    for line in lines:
        l = line.lower()
        if "company" in l and not data["company"]:
            data["company"] = re.sub(r"(?i)company\s*[:\-]?\s*", "", line).strip()
        if "state" in l and not data["state"]:
            data["state"] = re.sub(r"(?i)state\s*[:\-]?\s*", "", line).strip()
        if "district" in l and not data["district"]:
            data["district"] = re.sub(r"(?i)district\s*[:\-]?\s*", "", line).strip()

    for line in lines:
        l = line.lower()
        if any(keyword in l for keyword in IGNORE_NAME_LINES):
            continue
        if any(x for x in [data["email"], data["phone"], data["company"], data["state"], data["district"]] if x and x.lower() in l):
            continue
        line_clean = re.sub(r"(?i)^(full\s*)?name\s*[:\-]?\s*", "", line).strip()
        clean_line = re.sub(r"[^A-Za-z.\-\' ]", "", line_clean).strip()
        words = clean_line.split()
        if not clean_line or len(words) > 5:
            continue
        if re.match(r"^[A-Za-z.\-\' ]+$", clean_line):
            if words and words[0].lower().rstrip('.') in NAME_PREFIXES:
                clean_line = " ".join(words[1:])
            data["name"] = clean_line.title()
            break
    return data

def extract_pdf(pdf_bytes, filename="Unknown"):
    try:
        pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return {"filename": filename, "name": "", "email": "", "phone": "", "company": "", "state": "", "district": "", "pin": ""}
    text = "\n".join([page.get_text() for page in pdf])
    pdf.close()
    data = extract_info_smart(text)
    data["filename"] = filename
    return data

# ===== Flask Routes =====
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    extracted_data = []
    for file in request.files.getlist("pdf"):
        filename = file.filename
        if filename.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(file.read())) as z:
                for inner_name in z.namelist():
                    if inner_name.lower().endswith(".pdf"):
                        with z.open(inner_name) as pdf_file:
                            data = extract_pdf(pdf_file.read(), inner_name)
                            extracted_data.append(data)
        elif filename.lower().endswith(".pdf"):
            data = extract_pdf(file.read(), filename)
            extracted_data.append(data)
        else:
            return jsonify({"error": "Only PDF or ZIP files are allowed!"})
    return jsonify({"autofill": extracted_data})

@app.route("/save", methods=["POST"])
def save_data():
    all_data = request.get_json()
    if not all_data:
        return jsonify({"message": "No data received!"})

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    table_name = f"extracted_data_{timestamp}"

    try:
        conn = mysql.connector.connect(
            host=DB_CONFIG["host"],
            user=DB_CONFIG["user"],
            password=DB_CONFIG["password"],
            database=DB_CONFIG["database"],
            auth_plugin='mysql_native_password'
        )
        cur = conn.cursor()

        cur.execute(f"""
            CREATE TABLE {table_name} (
                id INT AUTO_INCREMENT PRIMARY KEY,
                filename VARCHAR(255),
                name VARCHAR(255),
                email VARCHAR(255),
                phone VARCHAR(50),
                company VARCHAR(255),
                state VARCHAR(100),
                district VARCHAR(100),
                pin VARCHAR(20)
            )
        """)

        for entry in all_data:
            cur.execute(f"""
                INSERT INTO {table_name}
                (filename, name, email, phone, company, state, district, pin)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                entry.get("filename"), entry.get("name"), entry.get("email"),
                entry.get("phone"), entry.get("company"), entry.get("state"),
                entry.get("district"), entry.get("pin")
            ))

        conn.commit()
        cur.close()
        conn.close()

        with open("last_table.txt", "w") as f:
            f.write(table_name)

        return jsonify({"message": f"Saved {len(all_data)} record(s) to new table: {table_name}"})
    except mysql.connector.Error as e:
        return jsonify({"message": f"Database error: {str(e)}"})

@app.route("/download_excel")
def download_excel():
    try:
        if not os.path.exists("last_table.txt"):
            return "No saved table found. Please save data first!"

        with open("last_table.txt", "r") as f:
            table_name = f.read().strip()

        conn = mysql.connector.connect(
            host=DB_CONFIG["host"],
            user=DB_CONFIG["user"],
            password=DB_CONFIG["password"],
            database=DB_CONFIG["database"],
            auth_plugin='mysql_native_password'
        )
        cur = conn.cursor()
        cur.execute(f"SELECT filename, name, email, phone, company, state, district, pin FROM {table_name}")
        rows = cur.fetchall()

        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Data"
        headers = ["Filename", "Name", "Email", "Phone", "Company", "State", "District", "Pin"]
        ws.append(headers)
        for row in rows:
            ws.append(row)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        cur.close()
        conn.close()

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_name = f"Extracted_Data_{timestamp}.xlsx"

        return send_file(output, as_attachment=True, download_name=excel_name,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return f"Error generating Excel: {str(e)}"

if __name__ == "__main__":
    app.run(debug=True)
