import datetime
import hashlib
import json
import os
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Local Conveyance App", page_icon="🚗", layout="wide"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "Local Conveyance Template.xlsx")
CREDENTIALS_FILE = os.path.join(BASE_DIR, "users_credentials.json")

# --- MOBILE RESPONSIVE CSS ---
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    </style>
""",
    unsafe_allow_html=True,
)


def hash_password(password):
  return hashlib.sha256(password.encode()).hexdigest()


def load_users():
  if os.path.exists(CREDENTIALS_FILE):
    try:
      with open(CREDENTIALS_FILE, "r") as f:
        return json.load(f)
    except Exception:
      return {}
  return {}


def save_users(users_data):
  with open(CREDENTIALS_FILE, "w") as f:
    json.dump(users_data, f, indent=4)


def convert_number_to_words(n):
  if n == 0:
    return "Zero"

  units = [
      "",
      "One",
      "Two",
      "Three",
      "Four",
      "Five",
      "Six",
      "Seven",
      "Eight",
      "Nine",
      "Ten",
      "Eleven",
      "Twelve",
      "Thirteen",
      "Fourteen",
      "Fifteen",
      "Sixteen",
      "Seventeen",
      "Eighteen",
      "Nineteen",
  ]
  tens = [
      "",
      "",
      "Twenty",
      "Thirty",
      "Forty",
      "Fifty",
      "Sixty",
      "Seventy",
      "Eighty",
      "Ninety",
  ]

  def convert_below_thousand(num):
    if num == 0:
      return ""
    elif num < 20:
      return units[num]
    elif num < 100:
      return (
          tens[num // 10]
          + (" " + units[num % 10] if num % 10 != 0 else "")
      )
    else:
      return (
          units[num // 100]
          + " Hundred"
          + (
              " and " + convert_below_thousand(num % 100)
              if num % 100 != 0
              else ""
          )
      )

  n = round(n, 2)
  integer_part = int(n)
  fractional_part = int(round((n - integer_part) * 100))

  if integer_part == 0:
    word_str = "Zero"
  else:
    crore = integer_part // 10000000
    lakh = (integer_part // 100000) % 100
    thousand = (integer_part // 1000) % 100
    hundred = (integer_part // 100) % 10
    remainder = integer_part % 100

    res = []
    if crore > 0:
      res.append(convert_below_thousand(crore) + " Crore")
    if lakh > 0:
      res.append(convert_below_thousand(lakh) + " Lakh")
    if thousand > 0:
      res.append(convert_below_thousand(thousand) + " Thousand")
    if hundred > 0:
      res.append(units[hundred] + " Hundred")
    if remainder > 0:
      if crore > 0 or lakh > 0 or thousand > 0 or hundred > 0:
        res.append("and " + convert_below_thousand(remainder))
      else:
        res.append(convert_below_thousand(remainder))
    word_str = " ".join(res)

  result = f"Rupees {word_str}"
  if fractional_part > 0:
    result += f" and {convert_below_thousand(fractional_part)} Paise"
  return result + " Only"


# --- INITIALIZE SESSION STATE & AUTO-LOGIN VIA QUERY PARAMS ---
if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
if "username" not in st.session_state:
  st.session_state.username = ""

users_db = load_users()

# Check query params for persistent auto-login
if not st.session_state.logged_in and "session_user" in st.query_params:
  saved_user = st.query_params["session_user"]
  if saved_user in users_db:
    st.session_state.logged_in = True
    st.session_state.username = saved_user

if "records" not in st.session_state:
  st.session_state.records = []
if "form_version" not in st.session_state:
  st.session_state.form_version = 0

# Load user-specific records and profile data if logged in
if st.session_state.logged_in:
  user_data = users_db.get(st.session_state.username, {})
  st.session_state.records = user_data.get("records", [])
  user_profile = user_data.get("profile", {})
  user_bank = user_data.get("bank", {})
else:
  user_profile = {}
  user_bank = {}


def save_user_data_to_db():
  db = load_users()
  if st.session_state.username in db:
    db[st.session_state.username]["records"] = st.session_state.records
    db[st.session_state.username]["profile"] = st.session_state.get(
        "current_profile", {}
    )
    db[st.session_state.username]["bank"] = st.session_state.get(
        "current_bank", {}
    )
    save_users(db)


# --- MAIN SCREEN: AUTHENTICATION (IF NOT LOGGED IN) ---
if not st.session_state.logged_in:
  st.title("🚗 Local Conveyance Tracker & Excel Generator")
  st.markdown(
      "### Welcome! Please log in or register to manage your conveyance claims"
      " securely."
  )

  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    auth_tab = st.radio("Choose Action", ["Login", "Sign Up"], horizontal=True)
    input_user = st.text_input("Username / Email", placeholder="Enter username")
    input_pass = st.text_input(
        "Password", type="password", placeholder="Enter password"
    )

    if auth_tab == "Sign Up":
      if st.button("Register Account", type="primary", use_container_width=True):
        clean_user = input_user.strip().lower()
        if not clean_user or not input_pass:
          st.error("Please fill in both fields.")
        elif clean_user in users_db:
          st.error("Username already exists! Please login instead.")
        else:
          users_db[clean_user] = {
              "password": hash_password(input_pass),
              "records": [],
              "profile": {},
              "bank": {},
          }
          save_users(users_db)
          st.success("Account created successfully! You can now log in.")
    else:
      if st.button("Login", type="primary", use_container_width=True):
        clean_user = input_user.strip().lower()
        if clean_user in users_db and users_db[clean_user]["password"] == hash_password(
            input_pass
        ):
          st.session_state.logged_in = True
          st.session_state.username = clean_user
          st.query_params["session_user"] = clean_user
          st.success("Logged in successfully!")
          st.rerun()
        else:
          st.error("Invalid username or password.")
  st.stop()  # Stop rendering dashboard until logged in


# --- DASHBOARD HEADER & TABS (WHEN LOGGED IN) ---
st.title("🚗 Local Conveyance Tracker & Excel Generator")
st.markdown(f"Logged in user: **{st.session_state.username}**")

tab_entries, tab_profile, tab_account = st.tabs(
    ["🚗 Travel Entries & Export", "👤 Profile & Bank Details", "⚙️ Account"]
)

# --- TAB 2: PROFILE & BANK DETAILS (PERMANENTLY SAVED) ---
with tab_profile:
  st.subheader("👤 Employee & Banking Information")
  st.markdown(
      "Values entered here are automatically saved to your profile and will"
      " persist across reloads."
  )

  with st.form("profile_bank_form"):
    col_p1, col_p2 = st.columns(2)
    with col_p1:
      st.markdown("#### Employee Profile")
      emp_name = st.text_input(
          "Employee Name", value=user_profile.get("emp_name", "")
      )
      emp_code = st.text_input(
          "Employee Code", value=user_profile.get("emp_code", "")
      )
      entity_name = st.text_input(
          "Entity / Company Name", value=user_profile.get("entity_name", "")
      )
      branch = st.text_input(
          "Branch / Location", value=user_profile.get("branch", "")
      )
      lob = st.text_input("LOB / Sector", value=user_profile.get("lob", ""))
      period = st.text_input(
          "Period (Month/Year)", value=user_profile.get("period", "")
      )

    with col_p2:
      st.markdown("#### Bank & Payment Details")
      bank_name = st.text_input(
          "Bank Name", value=user_bank.get("bank_name", "")
      )
      bank_acc = st.text_input(
          "Bank Account Number", value=user_bank.get("bank_acc", "")
      )
      ifsc_code = st.text_input(
          "IFSC Code", value=user_bank.get("ifsc_code", "")
      )
      pan_number = st.text_input(
          "PAN Number", value=user_bank.get("pan_number", "")
      )

    submitted_profile = st.form_submit_button(
        "💾 Save Profile & Bank Details", type="primary"
    )
    if submitted_profile:
      st.session_state["current_profile"] = {
          "emp_name": emp_name,
          "emp_code": emp_code,
          "entity_name": entity_name,
          "branch": branch,
          "lob": lob,
          "period": period,
      }
      st.session_state["current_bank"] = {
          "bank_name": bank_name,
          "bank_acc": bank_acc,
          "ifsc_code": ifsc_code,
          "pan_number": pan_number,
      }
      save_user_data_to_db()
      st.success("Profile and banking details saved permanently!")
      st.rerun()

# Sync current active values for Excel generation
current_profile = user_profile
current_bank = user_bank


# --- TAB 1: TRAVEL ENTRIES & EXCEL GENERATOR ---
with tab_entries:
  st.subheader("➕ Add New Travel Entry")
  fv = st.session_state.form_version

  col1, col2, col3 = st.columns(3)

  with col1:
    travel_date = st.date_input("Date", key=f"input_date_{fv}")
    from_loc = st.text_input(
        "From Location", key=f"input_from_{fv}", placeholder="e.g., Office"
    )
    to_loc = st.text_input(
        "To Location", key=f"input_to_{fv}", placeholder="e.g., Client Site"
    )

  with col2:
    reason_choice = st.selectbox(
        "Particulars / Purpose",
        ["Meeting", "Other"],
        index=0,
        key=f"input_reason_choice_{fv}",
    )
    if reason_choice == "Other":
      custom_purpose = st.text_input(
          "Specify Custom Reason",
          key=f"input_custom_purpose_{fv}",
          placeholder="Type reason...",
      )
      purpose = custom_purpose
    else:
      purpose = "Meeting"

  with col3:
    distance = st.number_input(
        "Kilometres (if personal vehicle)",
        min_value=0.0,
        step=1.0,
        value=None,
        format="%g",
        placeholder="Type value...",
        key=f"input_dist_{fv}",
    )
    rate_km = st.number_input(
        "Rate per KM (INR)",
        min_value=0.0,
        step=1.0,
        value=None,
        format="%g",
        placeholder="Type value...",
        key=f"input_rate_{fv}",
    )
    manual_amount = st.number_input(
        "Amount Used / Direct Amount (INR)",
        min_value=0.0,
        step=1.0,
        value=None,
        format="%g",
        placeholder="Type value...",
        key=f"input_amount_{fv}",
    )

  if st.button("💾 Save Entry", type="primary"):
    dist_val = distance if distance is not None else 0.0
    r_val = rate_km if rate_km is not None else 0.0
    m_amt = manual_amount if manual_amount is not None else 0.0

    calc_amount = m_amt if m_amt > 0 else (dist_val * r_val)

    new_entry = {
        "id": str(datetime.datetime.now().timestamp()),
        "Date": travel_date.strftime("%d/%m/%Y"),
        "Particulars": purpose,
        "From Location": from_loc,
        "To Location": to_loc,
        "Kilometres": dist_val,
        "Rate/KM": r_val,
        "Amount": calc_amount,
    }

    st.session_state.records.append(new_entry)
    save_user_data_to_db()
    st.session_state.form_version += 1
    st.success("Entry saved successfully!")
    st.rerun()

  st.divider()
  st.subheader("📋 Past Days' Records")

  if st.session_state.records:
    df_records = pd.DataFrame(st.session_state.records)

    total_amount_live = (
        df_records["Amount"].sum() if "Amount" in df_records.columns else 0.0
    )
    st.metric(
        label="💰 Simultaneous Total Claim Amount",
        value=f"₹{total_amount_live:,.2f}",
    )

    st.markdown(
        "💡 *Click directly on any cell in the table below to edit locations,"
        " amounts, or particulars right then and there!*"
    )

    display_df = df_records.drop(columns=["id"])

    edited_df = st.data_editor(
        display_df,
        use_container_width=True,
        key="editable_conveyance_table",
        num_rows="fixed",
        column_config={
            "Particulars": st.column_config.TextColumn(
                "Particulars / Purpose", required=True
            ),
            "From Location": st.column_config.TextColumn(
                "From Location", required=True
            ),
            "To Location": st.column_config.TextColumn(
                "To Location", required=True
            ),
            "Kilometres": st.column_config.NumberColumn(
                "Kilometres", min_value=0.0, format="%g"
            ),
            "Rate/KM": st.column_config.NumberColumn(
                "Rate/KM", min_value=0.0, format="₹%g"
            ),
            "Amount": st.column_config.NumberColumn(
                "Amount", min_value=0.0, format="₹%g"
            ),
        },
    )

    updated_records = []
    for idx, row in edited_df.iterrows():
      original_entry = st.session_state.records[idx]
      updated_entry = {
          "id": original_entry["id"],
          "Date": str(row["Date"]),
          "Particulars": str(row["Particulars"]),
          "From Location": str(row["From Location"]),
          "To Location": str(row["To Location"]),
          "Kilometres": (
              float(row["Kilometres"]) if pd.notna(row["Kilometres"]) else 0.0
          ),
          "Rate/KM": float(row["Rate/KM"]) if pd.notna(row["Rate/KM"]) else 0.0,
          "Amount": float(row["Amount"]) if pd.notna(row["Amount"]) else 0.0,
      }
      updated_records.append(updated_entry)

    if updated_records != st.session_state.records:
      st.session_state.records = updated_records
      save_user_data_to_db()

    selected_row = st.selectbox(
        "Select an entry to delete if needed:",
        options=range(len(st.session_state.records)),
        format_func=lambda x: (
            f"{st.session_state.records[x]['Date']} |"
            f" {st.session_state.records[x]['From Location']} ➔"
            f" {st.session_state.records[x]['To Location']}"
            f" (₹{st.session_state.records[x]['Amount']})"
        ),
    )

    if st.button("❌ Delete Selected Entry"):
      removed = st.session_state.records.pop(selected_row)
      save_user_data_to_db()
      st.success(
          f"Deleted entry from {removed['Date']} ({removed['From Location']} ➔"
          f" {removed['To Location']})"
      )
      st.rerun()


    # --- DIRECT TEMPLATE FILLING FUNCTION ---
    def generate_styled_excel():
      if not os.path.exists(TEMPLATE_PATH):
        st.error(f"Error: Template not found at {TEMPLATE_PATH}")
        return None

      wb = openpyxl.load_workbook(TEMPLATE_PATH)
      ws = wb["June 26"]

      # Populate Profile Data
      ws["B4"] = current_profile.get("emp_name", "")
      ws["B6"] = current_profile.get("emp_code", "")
      ws["B8"] = current_profile.get("entity_name", "")
      ws["B10"] = current_profile.get("branch", "")
      ws["B12"] = current_profile.get("lob", "")
      ws["B14"] = current_profile.get("period", "")

      for r in range(17, 50):
        for c in range(1, 8):
          ws.cell(row=r, column=c).value = None

      current_row = 17
      total_amount = 0.0

      for rec in st.session_state.records:
        if current_row >= 50:
          break

        row_data = [
            rec["Date"],
            rec["Particulars"],
            rec["From Location"],
            rec["To Location"],
            rec["Kilometres"],
            rec["Rate/KM"],
            rec["Amount"],
        ]
        total_amount += float(rec["Amount"])

        for col_num, val in enumerate(row_data, 1):
          cell = ws.cell(row=current_row, column=col_num, value=val)
          if col_num == 1:
            cell.alignment = Alignment(horizontal="center", vertical="center")
          if col_num == 5:
            cell.number_format = "#,##0.00"
          elif col_num in [6, 7]:
            cell.number_format = "₹#,##0.00"

        current_row += 1

      ws["C50"] = "Total Amount"
      ws["G50"] = "=SUM(G17:G49)"
      ws["G50"].number_format = "₹#,##0.00"

      words_str = convert_number_to_words(total_amount)
      ws["A51"] = "Please reimburse Rupees (in words):"
      ws["C51"] = words_str

      # Populate Bank Details
      ws["B55"] = current_bank.get("bank_name", "")
      ws["B56"] = current_bank.get("bank_acc", "")

      # Signatures
      ws["B61"] = current_profile.get("emp_name", "")
      ws["B62"] = datetime.date.today().strftime("%d/%m/%Y")

      output_path = os.path.join(BASE_DIR, "Filled_Local_Conveyance.xlsx")
      wb.save(output_path)
      return output_path


    # --- EXCEL DOWNLOAD SECTION ---
    st.divider()
    st.subheader("📥 Export Styled Excel Sheet")

    if st.button("🔄 Generate Fresh Excel Sheet", type="primary"):
      st.session_state.excel_path = generate_styled_excel()
      if st.session_state.excel_path:
        st.success("Excel sheet generated successfully from template!")

    if "excel_path" in st.session_state and os.path.exists(
        st.session_state.excel_path
    ):
      with open(st.session_state.excel_path, "rb") as f:
        st.download_button(
            label="⬇️ Click Here to Download Generated Excel",
            data=f,
            file_name=(
                f"Conveyance_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
  else:
    st.info("No travel records found yet. Fill out the form above to get started!")

# --- TAB 3: ACCOUNT & SESSION CONTROLS ---
with tab_account:
  st.subheader("⚙️ Account Controls")
  if st.button("🚪 Log Out", type="secondary"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.records = []
    if "session_user" in st.query_params:
      del st.query_params["session_user"]
    st.rerun()

  st.divider()
  if st.button("🗑️ Clear All Saved Travel Entries", type="primary"):
    st.session_state.records = []
    save_user_data_to_db()
    st.success("All travel records cleared.")
    st.rerun()
