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
with st.spinner(text="Reading data from spreadsheet..."):
    # pbar = st.progress(20, text="Reading data from spreadsheet...")
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

    # st.write(list(zip(exam_names, exam_demands)))
    # pbar.progress(20)

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

    # st.write([(date, capacity) for date, capacity in zip(dates, dates_capacity)])
    # pbar.progress(40)

    # Extract gap constraints
    min_days_between_exams = {}
    for row in data_rows:
        exam1, exam2, min_days = row[9], row[10], row[11]
        if exam1 and exam2 and min_days:
            exam1 = exam_index[exam1]
            exam2 = exam_index[exam2]
            min_days = int(min_days)
            min_days_between_exams[(exam1, exam2)] = min_days

    # st.write(min_days_between_exams)
    # pbar.progress(60)

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

    # st.write(exam_before_exam)
    # st.write(exam_before_date)
    # pbar.progress(80)

    # Extract prescheduled constraints
    exam_on_date = []
    for row in data_rows:
        exam, date = row[22], row[23]
        if exam and date:
            exam = exam_index[exam]
            date = date_index[date]
            exam_on_date.append((exam, date))

    # st.write(exam_on_date)
    # pbar.progress(100)



#### Solve scheduling problem ####

with st.spinner(text="Solving scheduling problem..."):
    from ortools.sat.python import cp_model
    from datetime import datetime

    # Define the number of exams and the number of days
    num_exams = len(exam_names)
    horizon = len(dates)

    # Create a CP-SAT model
    model = cp_model.CpModel()

    # Create variables
    exams = [model.NewIntVar(1, horizon, f'exam_{i}') for i in range(num_exams)]

    # Add minimal gap constraints
    for (i, j), min_days in min_days_between_exams.items():
        # Interval for each exam
        interval_i = model.NewFixedSizeIntervalVar(exams[i], min_days, f'task_{i,j}')
        interval_j = model.NewFixedSizeIntervalVar(exams[j], min_days, f'task_{j,i}')
        model.AddNoOverlap([interval_i, interval_j])
        # model.Add(exams[i] + min_days <= exams[j] or exams[j] + min_days <= exams[i])

    # Add daily capacity constraints
    max_capacity = max(dates_capacity)
    exam_intervals = [model.NewFixedSizeIntervalVar(exams[i], 1, f'demand_{i}') for i in range(num_exams)]
    fake_intervals = [model.NewFixedSizeIntervalVar(t, 1, f'fake_demand_{t}') for t in range(horizon)]
    all_intervals = exam_intervals + fake_intervals
    all_demands = exam_demands + [max_capacity - c for c in dates_capacity]
    model.AddCumulative(all_intervals, all_demands, max_capacity)

    # Add precedence constraints
    for (i,j) in exam_before_exam:
        model.Add(exams[i] < exams[j])
    for (i,t) in exam_before_date:
        model.Add(exams[i] < t)

    # Add prescheduling constraints
    for (i,t) in exam_on_date:
        model.Add(exams[i] == t)


    # Define the objective
    makespan = model.NewIntVar(0, horizon, 'makespan')
    model.AddMaxEquality(makespan, exams)
    model.Minimize(makespan)

    # Create a solver and solve the model
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    # dump solution into a dictionary
    solution = {}
    if status == cp_model.OPTIMAL:
        print('Exam schedule:')
        for i in range(num_exams):
            exam = exam_names[i]
            date = dates[solver.Value(exams[i])]
            date = datetime.strptime(date, '%d/%m/%Y').date()
            solution[exam] = date

    # Print the solution
    if len(solution) > 0:
        sorted_items = sorted(solution.items(), key=lambda x:x[1])
        for i, (exam, date) in enumerate(sorted_items):
            datestr = date.strftime('%d/%m/%Y')
            # datestr = date
            st.write(f'{datestr} : {exam}')
    else:
        st.write('No solution found')