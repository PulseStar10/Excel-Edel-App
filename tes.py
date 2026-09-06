import datetime
import json
import os
import re
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


# --- SIDEBAR: LOGIN & PROFILE ---
with st.sidebar:
  st.header("🔐 Secure Access")
  login_id = st.text_input(
      "Your Private ID / Username",
      value="",
      placeholder="e.g., john_doe_123",
      help=(
          "Type your unique ID here. Your data permanently binds to this ID"
          " across reloads!"
      ),
  )

  # Sanitize login id for filename usage
  safe_id = (
      re.sub(r"[^a-zA-Z0-9_-]", "", login_id.strip().lower())
      if login_id
      else ""
  )
  DATA_FILE = (
      os.path.join(BASE_DIR, f"conveyance_store_{safe_id}.json")
      if safe_id
      else ""
  )


  def load_data():
    if safe_id and os.path.exists(DATA_FILE):
      try:
        with open(DATA_FILE, "r") as f:
          return json.load(f)
      except Exception:
        return []
    return []


  def save_data(data):
    if safe_id:
      with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=4)


  st.divider()
  st.header("👤 Employee Profile")
  emp_name = st.text_input("Employee Name", value="", placeholder="e.g., ABC XYZ")
  emp_code = st.text_input(
      "Employee Code", value="", placeholder="e.g., 6-digit No."
  )
  entity_name = st.text_input(
      "Entity / Company Name", value="", placeholder="e.g., Company Name"
  )
  branch = st.text_input(
      "Branch / Location", value="", placeholder="e.g., City District"
  )
  lob = st.text_input("LOB / Sector", value="", placeholder="e.g., Sector Name")
  period = st.text_input(
      "Period (Month/Year)", value="", placeholder="e.g., June 2026"
  )

  st.divider()
  st.header("🏦 Bank & Payment Details")
  bank_name = st.text_input("Bank Name", value="", placeholder="e.g., HDFC Bank")
  bank_acc = st.text_input(
      "Bank Account Number", value="", placeholder="Account Number"
  )
  ifsc_code = st.text_input("IFSC Code", value="", placeholder="IFSC Code")
  pan_number = st.text_input("PAN Number", value="", placeholder="PAN Number")

  st.divider()
  st.markdown("### 📊 Data Management")
  if st.button("🗑️ Clear My Saved Data", type="secondary"):
    if safe_id and os.path.exists(DATA_FILE):
      os.remove(DATA_FILE)
    st.success("Your saved data cleared!")
    st.rerun()

st.title("🚗 Local Conveyance Tracker & Excel Generator")

# Check if user logged in
if not safe_id:
  st.warning(
      "⚠️ **Please enter your Private ID / Username in the sidebar to load"
      " and save your secure records.**"
  )
else:
  # Initialize records from file storage
  if "current_user_loaded" not in st.session_state or st.session_state.get(
      "active_user"
  ) != safe_id:
    st.session_state.records = load_data()
    st.session_state.active_user = safe_id
    st.session_state.current_user_loaded = True

  if "form_version" not in st.session_state:
    st.session_state.form_version = 0

  st.markdown(
      f"Welcome back, **{login_id}**! Your data is permanently linked to your"
      " ID and will never disappear on reload."
  )

  # --- MAIN FORM: ADD NEW ENTRY ---
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
    save_data(st.session_state.records)
    st.session_state.form_version += 1
    st.success("Entry saved successfully!")
    st.rerun()

  # --- DISPLAY & INLINE EDITABLE RECORDS & TOTAL ---
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
      save_data(st.session_state.records)

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
      save_data(st.session_state.records)
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

      ws["B4"] = emp_name
      ws["B6"] = emp_code
      ws["B8"] = entity_name
      ws["B10"] = branch
      ws["B12"] = lob
      ws["B14"] = period

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

      ws["B55"] = bank_name
      ws["B56"] = bank_acc

      ws["B61"] = emp_name
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
    st.info(
        "No travel records found for this ID yet. Fill out the form above to"
        " get started!"
    )
