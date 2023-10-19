import streamlit as st
import gspread
import re
from google.oauth2 import service_account

#### function defs ####
def get_matching(pattern, names, index):
    return [index.get(name) for name in names if re.match(pattern1,name)]
    # return [i for i in matches if i is not None]

def get_matching_pairs(pattern1, pattern2, names, index):
    matches = [name for name in names if re.match(pattern1,name)]
    subs = [re.sub(pattern1,pattern2,name) for name in matches]

    pairs = []
    for match,sub in zip(matches,subs):
        i1 = index.get(match)
        i2 = index.get(sub)
        if i1 and i2 and i1 != i2: pairs.append((i1,i2))
    return pairs


#### Authorize and connect to Sheets ####
credentials = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
    ],
)
gc = gspread.authorize(credentials)


#### Hello ####

st.title('Exam scheduler 2024')

st.write('Enter data in spreadsheet:')
st.write(st.secrets["private_gsheets_url"])
if not st.button("Process!"):
    st.stop()


#### Read Google Sheets input ####
message = "reading data from spreadsheet"
with st.spinner(text=message.capitalize() + '...'):
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
        name, demand = row[1].strip(), row[2].strip()
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
        date, capacity = row[5].strip(), row[6].strip()
        if date:
            capacity = int(capacity) if capacity else 0
            date_index[date] = len(dates)
            dates.append(date)
            dates_capacity.append(capacity)

    # st.write([(date, capacity) for date, capacity in zip(dates, dates_capacity)])
    # pbar.progress(40)

    # Extract minimal and ideal gap constraints
    min_days_between_exams = {}
    ideal_days_between_exams = {}
    for row in data_rows:
        pattern1, pattern2, min_days, ideal_days = row[9].strip(), row[10].strip(), row[11].strip(), row[12].strip()

        if not (pattern1 and pattern2): continue
        pairs = get_matching_pairs(pattern1,pattern2,exam_names,exam_index)

        if min_days:
            min_days = int(min_days)
            for (exam1, exam2) in pairs:
                min_days_between_exams[(exam1, exam2)] = min_days
        
        if ideal_days:
            ideal_days = int(ideal_days)
            for (exam1, exam2) in pairs:
                ideal_days_between_exams[(exam1, exam2)] = ideal_days


    # st.write(min_days_between_exams)
    # pbar.progress(60)

    # Extract precedence constraints
    exam_before_exam = []
    for row in data_rows:
        pattern1, pattern2 = row[15].strip(), row[16].strip()
        if not (pattern1 and pattern2): continue
        
        pairs = get_matching_pairs(pattern1,pattern2,exam_names,exam_index)
        for (exam1, exam2) in pairs:
            exam_before_exam.append((exam1, exam2))

    # exam_before_date = []
    # for row in data_rows:
    #     exam, date = row[18], row[19]
    #     if exam and date:
    #         exam = exam_index[exam]
    #         date = date_index[date]
    #         exam_before_date.append((exam, date))

    # st.write(exam_before_exam)
    # st.write(exam_before_date)
    # pbar.progress(80)

    # Extract prescheduled constraints
    exam_on_date = []
    for row in data_rows:
        pattern, date = row[19].strip(), row[20].strip()
        if not (exam and date): continue

        matches = get_matching(pattern,exam_names,exam_index)
        date = date_index[date]

        for exam in matches: 
            exam_on_date.append((exam, date))

    # st.write(exam_on_date)
    # pbar.progress(100)


st.success('Done ' + message)


#### Solve scheduling problem ####
message = "solving scheduling problem"
with st.spinner(text=message.capitalize() + '...'):
    from ortools.sat.python import cp_model
    from datetime import datetime

    # Define the number of exams and the number of days
    num_exams = len(exam_names)
    horizon = len(dates)

    # Create a CP-SAT model
    model = cp_model.CpModel()

    # Create variables
    exams = [model.NewIntVar(0, horizon-1, f'exam_{i}') for i in range(num_exams)]

    # Add minimal gap constraints
    for (i, j), days in min_days_between_exams.items():
        # Interval for each exam
        interval_i = model.NewFixedSizeIntervalVar(exams[i], days, f'mingap_{i,j}')
        interval_j = model.NewFixedSizeIntervalVar(exams[j], days, f'mingap_{j,i}')
        model.AddNoOverlap([interval_i, interval_j])
        # model.Add(exams[i] + min_days <= exams[j] or exams[j] + min_days <= exams[i])

    # Add ideal gap constraints
    ideal_bools = {}
    for (i, j), days in ideal_days_between_exams.items():
        b = model.NewBoolVar(f'idealbool_{i,j}')
        ideal_bools[(i,j)] = b

        # Interval for each exam
        interval_i = model.NewOptionalFixedSizeIntervalVar(exams[i], days, b, f'idealgap_{i,j}')
        interval_j = model.NewOptionalFixedSizeIntervalVar(exams[j], days, b, f'idealgap_{j,i}')
        model.AddNoOverlap([interval_i, interval_j])

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
    # for (i,t) in exam_before_date:
    #     model.Add(exams[i] < t)

    # Add prescheduling constraints
    for (i,t) in exam_on_date:
        model.Add(exams[i] == t)


    # Define the objective
    collisions = []
    for i in range(num_exams):
        for j in range(num_exams):
            b = model.NewBoolVar(f'{i}{j}')
            model.Add(exams[i]==exams[j]).OnlyEnforceIf(b)
            model.Add(exams[i]!=exams[j]).OnlyEnforceIf(b.Not())
            collisions.append(b)

    factor = num_exams**2
    if len(ideal_bools) > 0:
        # Minimize collisions, but prioritize soft constraints
        model.Minimize( -factor * sum(ideal_bools.values()) + sum(collisions) )
    else:
        # Minimize collisions
        model.Minimize( sum(collisions) )

    # Define the objective: makespan
    # makespan = model.NewIntVar(0, horizon, 'makespan')
    # model.AddMaxEquality(makespan, exams)
    # model.Minimize(makespan)

    # Create a solver and solve the model
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    # check status
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        st.error('No solution found :(')
        st.stop()

    # Solution found!
    st.balloons()
    message = 'an optimal' if status == cp_model.OPTIMAL else 'a feasible'
    st.success(f'Found {message} solution!')

    for (i,j),b in ideal_bools.items():
        if not solver.Value(b):
            st.warning(f'could not satisfy ideal gap: {exam_names[i]}, {exam_names[j]}', icon="⚠️")

    # dump solution into a dictionary
    solution = {}
    for i in range(num_exams):
        exam = exam_names[i]
        date = dates[solver.Value(exams[i])]
        date = datetime.strptime(date, '%d/%m/%Y').date()
        solution[exam] = date

    # # Print the solution
    # if len(solution) > 0:
    #     sorted_items = sorted(solution.items(), key=lambda x:x[1])
    #     st.write('Exam schedule:')
    #     for i, (exam, date) in enumerate(sorted_items):
    #         datestr = date.strftime('%d/%m/%Y')
    #         # datestr = date
    #         st.write(f'{datestr} : {exam}')
    # else:
    #     st.write('No solution found')


#### Save solution to the Google Sheet ####
message = "writing output to spreadsheet"
with st.spinner(text=message.capitalize() + '...'):
    # Open the 'Output' worksheet
    output = workbook.worksheet('Output')

    # Clear existing content in the 'Output' worksheet starting from row 3
    start_row = 3
    end_row = output.row_count
    output.batch_clear([f'B{start_row}:C{end_row}'])

    if len(solution) > 0:
        sorted_items = sorted(solution.items(), key=lambda x: x[1])
        for i, (exam, date) in enumerate(sorted_items):
            date = date.strftime('%d/%m/%Y')
            output.append_row([exam, date], value_input_option="USER_ENTERED")
        
        # Style dates in column C
        date_format = {'numberFormat': {'type': 'DATE', 'pattern': 'dd/mm/yyyy'}}
        date_range = 'C3:C' + str(len(sorted_items) + 2)  # Range excluding header row
        output.format(date_range, date_format)

        st.success('Done ' + message)
    else:
        output.append_row(['', 'No solution found :('])
