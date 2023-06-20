import streamlit as st
import gspread
from google.oauth2 import service_account

#### Authorize and connect to Sheets ####
credentials = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
    ],
)
gc = gspread.authorize(credentials)


#### Read Google Sheets input ####
st.write("Reading data from spreadsheet...")
sheet_url = st.secrets["private_gsheets_url"]
workbook = gc.open_by_url(sheet_url)
worksheet = workbook.worksheet('Input')
data_rows = worksheet.get_all_values()[2:]  # Exclude header rows

# Extract exams
exam_names = []
exam_demands = []
exam_index = {}
for row in data_rows:
    name, demand = row[1], row[2]
    if name:
        exam_index[name] = len(exam_names)
        demand = int(demand)
        exam_names.append(name)
        exam_demands.append(demand)

st.write(list(zip(exam_names, exam_demands)))

# Extract dates
dates = []
dates_capacity = []
date_index = {}
for row in data_rows:
    date, capacity = row[5], row[6]
    if date:
        capacity = int(capacity) if capacity else 0
        date_index[date] = len(dates)
        dates.append(date)
        dates_capacity.append(capacity)

# print([(date.strftime('%d/%m/%Y'), capacity) for date, capacity in zip(dates, dates_capacity)])
st.write([(date, capacity) for date, capacity in zip(dates, dates_capacity)])

# Extract gap constraints
min_days_between_exams = {}
for row in data_rows:
    exam1, exam2, min_days = row[9], row[10], row[11]
    if exam1 and exam2 and min_days:
        exam1 = exam_index[exam1]
        exam2 = exam_index[exam2]
        min_days = int(min_days)
        min_days_between_exams[(exam1, exam2)] = min_days

st.write(min_days_between_exams)

# Extract precedence constraints
exam_before_exam = []
for row in data_rows:
    exam1, exam2 = row[14], row[15]
    if exam1 and exam2:
        exam1 = exam_index[exam1]
        exam2 = exam_index[exam2]
        exam_before_exam.append((exam1, exam2))

exam_before_date = []
for row in data_rows:
    exam, date = row[18], row[19]
    if exam and date:
        exam = exam_index[exam]
        date = date_index[date]
        exam_before_date.append((exam, date))

st.write(exam_before_exam)
st.write(exam_before_date)

# Extract prescheduled constraints
exam_on_date = []
for row in data_rows:
    exam, date = row[22], row[23]
    if exam and date:
        exam = exam_index[exam]
        date = date_index[date]
        exam_on_date.append((exam, date))

st.write(exam_on_date)
