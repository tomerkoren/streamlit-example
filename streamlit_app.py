import streamlit as st
import gspread
import re
from openpyxl.utils.cell import get_column_letter
from ortools.sat.python import cp_model
from datetime import datetime
from google.oauth2 import service_account
import time

#### regex helper functions ####
def preprocess_name(name):
    # strip consecutive whitespaces
    name = re.sub(' +', ' ', name)
    return name

def preprocess_pattern(pattern):
    # strip consecutive whitespaces
    pattern = re.sub(' +', ' ', pattern)
    # use '#' as a wildcard character (in addition to '.')
    pattern = pattern.replace('#', '.')
    return pattern

def get_matching(pattern, names, index):
    return [index[name] for name in names if re.fullmatch(pattern,name)]

def get_matching_pairs(pattern1, pattern2, names, index):
    matches = get_matching(pattern1,names,index)
    subs = [re.sub(pattern1,pattern2,names[i]) for i in matches]

    pairs = []
    for i1,sub in zip(matches,subs):
        for i2 in get_matching(sub,names,index):
            if i1 != i2: pairs.append((i1,i2))
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

st.title('Exam scheduler 2024a')

st.write('Enter data in spreadsheet:')
st.write(st.secrets["private_gsheets_url"])

time_limit = st.slider('Time limit (seconds):', 1, 180, 20)

if not st.button("Process!"):
    st.stop()


#### Read Google Sheets input ####
message = "reading data from spreadsheet"
with st.spinner(text=message.capitalize() + '...'):
    log = st.expander("Log", expanded=False)

    # pbar = st.progress(20, text="Reading data from spreadsheet...")
    sheet_url = st.secrets["private_gsheets_url"]
    workbook = gc.open_by_url(sheet_url)

    # Extract exams
    worksheet = workbook.worksheet('בחינות')
    data_rows = worksheet.get_all_values()[2:]

    exam_names = []
    exam_demands = []
    exam_index = {}
    for row in data_rows:
        name, demand = row[1].strip(), row[2].strip()
        name = preprocess_name(name)
        if name:
            exam_index[name] = len(exam_names)
            demand = int(demand)
            exam_names.append(name)
            exam_demands.append(demand)

    # st.write(list(zip(exam_names, exam_demands)))
    # pbar.progress(20)

    # Extract dates
    worksheet = workbook.worksheet('תאריכים')
    data_rows = worksheet.get_all_values()[2:]

    dates = []
    dates_capacity = []
    date_index = {}
    for row_i, row in enumerate(data_rows):
        date, capacity = row[1].strip(), row[2].strip()
        if date:
            capacity = int(capacity) if capacity else 0
            date_index[date] = len(dates)
            dates.append(date)
            dates_capacity.append(capacity)

    # st.write([(date, capacity) for date, capacity in zip(dates, dates_capacity)])
    # pbar.progress(40)

    # Extract minimal and ideal gap constraints
    sheet_name = 'מרווחים'
    worksheet = workbook.worksheet(sheet_name)
    data_rows = worksheet.get_all_values()[2:]

    min_days_between_exams = {}
    ideal_days_between_exams = {}
    for row_i, row in enumerate(data_rows):
        pattern1, pattern2, min_days, ideal_days = row[1].strip(), row[2].strip(), row[3].strip(), row[4].strip()
        if not (pattern1 and pattern2): continue

        pattern1 = preprocess_pattern(pattern1)
        pattern2 = preprocess_pattern(pattern2)
        pairs = get_matching_pairs(pattern1,pattern2,exam_names,exam_index)
        if len(pairs) == 0:
            log.warning(f'Constraint in sheet {sheet_name}, row {row_i+3} yielded 0 matches', icon="⚠️")
        # st.write(f'found matching pairs for gap constraints: {pairs}')

        min_days = int(min_days) if min_days else None
        ideal_days = int(ideal_days) if ideal_days else None

        duplicates_found = False
        overriding = False
        for pair in pairs:
            # ensure that exam1 < exam2 to avoid duplicates
            if pair[0] == pair[1]: continue
            if pair[0] > pair[1]: (pair[0], pair[1]) = (pair[1], pair[0])

            # update values, detect duplicates
            if min_days:
                if pair in min_days_between_exams: 
                    duplicates_found = True
                    if min_days != min_days_between_exams[pair]:
                        overriding = True
                min_days_between_exams[pair] = min_days
            
            if ideal_days:
                if pair in ideal_days_between_exams:
                    duplicates_found = True
                    if ideal_days != ideal_days_between_exams[pair]:
                        overriding = True
                ideal_days_between_exams[pair] = ideal_days
        
        if duplicates_found:
            if overriding:
                log.warning(f'Duplicate constraint(s) detected in {sheet_name}, row {row_i+3} (OVERRIDING)', icon="⚠️")
            else:
                log.warning(f'Duplicate constraint(s) detected in {sheet_name}, row {row_i+3} (non-overriding)', icon="⚠️")

    # Filter out redundant constraints
    for (pair, min_days) in min_days_between_exams.items():
        ideal_days = ideal_days_between_exams.get(pair)
        if ideal_days and ideal_days <= min_days: 
            # disable constraint
            ideal_days_between_exams[pair] = 0

    # st.write(min_days_between_exams)
    # pbar.progress(60)

    # Extract precedence constraints
    sheet_name = 'קדימויות'
    worksheet = workbook.worksheet(sheet_name)
    data_rows = worksheet.get_all_values()[2:]

    exam_before_exam = []
    for row_i, row in enumerate(data_rows):
        pattern1, pattern2 = row[1].strip(), row[2].strip()
        if not (pattern1 and pattern2): continue
        
        pattern1 = preprocess_pattern(pattern1)
        pattern2 = preprocess_pattern(pattern2)
        pairs = get_matching_pairs(pattern1,pattern2,exam_names,exam_index)
        if len(pairs) == 0:
            log.warning(f'Constraint in sheet {sheet_name}, row {row_i+3} yielded 0 matches', icon="⚠️")
        # st.write(f'found {len(pairs)} matching pairs for precedence constraints')

        duplicates_found = False
        for (exam1, exam2) in pairs:
            # detect duplicates
            if (exam1, exam2) in exam_before_exam:
                duplicates_found = True
                exam_before_exam.remove((exam1, exam2))

            exam_before_exam.append((exam1, exam2))
        
        if duplicates_found:
            log.warning(f'Duplicate constraint(s) detected in {sheet_name}, row {row_i+3}', icon="⚠️")

    # exam_before_date = []
    # for row_i, row in enumerate(data_rows):
    #     exam, date = row[18], row[19]
    #     if exam and date:
    #         exam = exam_index[exam]
    #         date = date_index[date]
    #         exam_before_date.append((exam, date))

    # st.write(exam_before_exam)
    # st.write(exam_before_date)
    # pbar.progress(80)

    # Extract prescheduled constraints
    sheet_name = 'קיבועים'
    worksheet = workbook.worksheet(sheet_name)
    data_rows = worksheet.get_all_values()[2:]

    exam_on_date = []
    for row_i, row in enumerate(data_rows):
        pattern, date = row[1].strip(), row[2].strip()
        if not (pattern and date): continue

        pattern = preprocess_pattern(pattern)
        matches = get_matching(pattern,exam_names,exam_index)
        if len(matches) == 0:
            log.warning(f'Constraint in sheet {sheet_name}, row {row_i+3} yielded 0 matches', icon="⚠️")
        # st.write(f'found {len(matches)} matches for prescheduled constraints')
        date = date_index[date]

        duplicates_found = False
        for exam in matches: 
            # detect duplicates
            if (exam, date) in exam_on_date:
                duplicates_found = True
                exam_on_date.remove((exam, date))
            
            exam_on_date.append((exam, date))

        if duplicates_found:
            log.warning(f'Duplicate constraint(s) detected in {sheet_name}, row {row_i+3}', icon="⚠️")

st.success('Done ' + message)


#### Construct scheduling problem ####
# message = f'solving scheduling problem (limiting to {time_limit}s)'
# with st.progress(0, text=message.capitalize() + '...') as pbar:

# Define the number of exams and the number of days
num_exams = len(exam_names)
horizon = len(dates)

# Create a CP-SAT model
model = cp_model.CpModel()

# Create variables
exams = [model.NewIntVar(0, horizon-1, f'exam_{i}') for i in range(num_exams)]

# Add minimal gap constraints
for (i, j), days in min_days_between_exams.items():
    # ignore disabled constraints
    if days < 1: continue

    # Interval for each exam
    interval_i = model.NewFixedSizeIntervalVar(exams[i], days, f'mingap_{i,j}')
    interval_j = model.NewFixedSizeIntervalVar(exams[j], days, f'mingap_{j,i}')
    model.AddNoOverlap([interval_i, interval_j])
    # model.Add(exams[i] + min_days <= exams[j] or exams[j] + min_days <= exams[i])

# Add ideal gap constraints
ideal_bools = {}
for (i, j), days in ideal_days_between_exams.items():
    # ignore disabled constraints
    if days < 1: continue

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


# # Define the objective: minimize collisions
# collisions = []
# for i in range(num_exams):
#     for j in range(num_exams):
#         b = model.NewBoolVar(f'{i}{j}')
#         model.Add(exams[i]==exams[j]).OnlyEnforceIf(b)
#         model.Add(exams[i]!=exams[j]).OnlyEnforceIf(b.Not())
#         collisions.append(b)

# factor = num_exams**2
# if len(ideal_bools) > 0:
#     # Minimize collisions, but prioritize soft constraints
#     model.Minimize( -factor * sum(ideal_bools.values()) + sum(collisions) )
# else:
#     # Minimize collisions
#     model.Minimize( sum(collisions) )

# Define the objective: maximize soft constraints satisfaction
model.Maximize( sum(ideal_bools.values()) )

# # Define the objective: makespan
# makespan = model.NewIntVar(0, horizon, 'makespan')
# model.AddMaxEquality(makespan, exams)
# model.Minimize(makespan)

# Create a solver and solve the model
solver = cp_model.CpSolver()
# Sets a time limit
solver.parameters.max_time_in_seconds = time_limit

# Start progressbar
# message = f'solving scheduling problem (limiting to {time_limit}s)'
# pbar = st.progress(0, text=message.capitalize() + '...')
# st.session_state["counter"] = 0.0

# async def timer(pbar):
#     while True:
#         progress = st.session_state["counter"]
#         pbar.progress(progress, text=message.capitalize() + '...')
#         if progress == 1.0: return

#         incr = 1.0/time_limit
#         st.session_state["counter"] = min(progress+incr,1.0)
#         r = await asyncio.sleep(1)

# asyncio.run(timer(pbar))

# Solve!
message = f'solving scheduling problem (limiting to {time_limit}s)'
with st.spinner(text=message.capitalize() + '...'):
    status = solver.Solve(model)
    success = (status in [cp_model.OPTIMAL, cp_model.FEASIBLE])

# Complete progressbar
# st.session_state["counter"] = 1.0

if status == cp_model.UNKNOWN:
    st.error('No solution found within time limit :( Try increasing the limit.')
    st.stop()
elif status == cp_model.INFEASIBLE:
    st.error('The scheduling problem was proven infeasible :( Try relaxing some hard constraints.')
    st.stop()

# Solution found!
st.balloons()
message = 'an OPTIMAL' if status == cp_model.OPTIMAL else 'a FEASIBLE'
st.success(f'Found {message} solution')

# dump solution into a dictionary
solution = {}
for i in range(num_exams):
    exam = exam_names[i]
    date = dates[solver.Value(exams[i])]
    date = datetime.strptime(date, '%d/%m/%Y').date()
    solution[exam] = date

# dump failed soft constraints into a list
failed_list = []
for (i,j),b in ideal_bools.items():
    if not solver.Value(b):
        requested = ideal_days_between_exams[(i,j)]
        actual = abs(solver.Value(exams[i]) - solver.Value(exams[j]))
        failed_list.append((exam_names[i],exam_names[j],requested,actual))

if len(failed_list)>0:
    st.warning(f'Some requested gap constraints could not be satisfied (see output sheet)', icon="⚠️")

#### Save solution to the Google Sheet ####
message = "writing output to spreadsheet"
with st.spinner(text=message.capitalize() + '...'):
    # Open the 'Output' worksheet
    output = workbook.worksheet('שיבוץ')

    # Clear existing content in the 'Output' worksheet starting from row 3
    start_row = 3
    end_row = output.row_count
    output.batch_clear([f'B{start_row}:H{end_row}'])

    # dump solution into columns B:C
    sorted_items = sorted(solution.items(), key=lambda x: x[1])
    data = []
    for i, (exam, date) in enumerate(sorted_items):
        date = date.strftime('%d/%m/%Y')
        data.append([exam, date])
    # output.append_rows(data, value_input_option="USER_ENTERED")
    output.update(values=data, 
                  range_name=f'B{start_row}:C{start_row+len(data)-1}',
                  value_input_option="USER_ENTERED")
    
    # Style dates in column C
    date_format = {'numberFormat': {'type': 'DATE', 'pattern': 'dd/mm/yyyy'}}
    date_range = 'C3:C' + str(len(sorted_items) + 2)  # Range excluding header row
    output.format(date_range, date_format)

    # Dump failed soft constraints into columns E:H
    output.update(values=failed_list, 
                  range_name=f'E{start_row}:H{start_row+len(failed_list)-1}', 
                  value_input_option="USER_ENTERED")

st.success(f'All done!')