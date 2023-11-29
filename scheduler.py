import gspread
from ortools.sat.python import cp_model
from google.oauth2 import service_account
import re
import random
from datetime import datetime
from zoneinfo import ZoneInfo
import argparse
import tomllib
import csv

import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning) 

#### helper functions ####
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


def get_timestamp():
    timezone = ZoneInfo('Asia/Jerusalem')
    return datetime.now(tz=timezone).strftime("%d-%m-%Y %H:%M:%S")

# identify 'dummy' exams that should be omitted from output
def omit_from_output(exam_name):
    return exam_name.startswith('%')



def extract_solution_from_solver(solver, exam_vars, exam_names, dates):
    # dump solution into a dictionary
    solution = {}
    for i in range(num_exams):
        exam = exam_names[i]
        if omit_from_output(exam): continue
        date = dates[solver.Value(exam_vars[i])]
        solution[exam] = datetime.strptime(date, '%d/%m/%Y').date()
    
    return solution

# dump failed soft constraints into a list
def extract_violations_from_solver(solver, bool_vars_dict, exam_vars,
                                exam_names, requested_gaps):
    violations = []
    for (i,j),b in bool_vars_dict.items():
        if solver.Value(b):
            requested = requested_gaps[(i,j)]
            actual = abs(solver.Value(exam_vars[i]) - solver.Value(exam_vars[j]))
            violations.append((exam_names[i],exam_names[j],requested,actual))
    
    return violations

def write_solution_to_csv(fname, solution):
    # prepare solution
    sorted_items = sorted(solution.items(), key=lambda x: x[1])
    data = []
    for i, (exam, date) in enumerate(sorted_items):
        date = date.strftime('%d/%m/%Y')
        data.append([exam, date])

    with open(fname, 'w') as f:
        writer = csv.writer(f)
        writer.writerows(data)


def write_solution_to_gsheet(worksheet, solution, violations):
    # prepare solution
    sorted_items = sorted(solution.items(), key=lambda x: x[1])
    data = []
    for i, (exam, date) in enumerate(sorted_items):
        date = date.strftime('%d/%m/%Y')
        data.append([exam, date])

    # Clear existing content starting from row start_row
    start_row = 3
    end_row = worksheet.row_count
    worksheet.batch_clear([f'B{start_row}:H{end_row}'])
    
    # Write data into columns B:C
    range_name = f'B{start_row}:C{start_row+len(data)-1}'
    worksheet.update(range_name=range_name,
                  values=data, 
                  value_input_option="USER_ENTERED")
    
    # Style dates in column C
    date_format = {'numberFormat': {'type': 'DATE', 'pattern': 'dd/mm/yyyy'}}
    range_name = f'C{start_row}:C{start_row+len(data)-1}'  # Range excluding header row
    worksheet.format(range_name, date_format)

    # Dump failed soft constraints into columns E:H
    range_name = f'E{start_row}:H{start_row+len(violations)-1}'
    worksheet.update(range_name=range_name, 
                    values=violations, 
                    value_input_option="USER_ENTERED")


# simple logger
logger = []
def log(str):
    str = f'{get_timestamp()} >>> {str}'
    print(str)
    logger.append(str)

# Solver callback
class MySolutionCallback(cp_model.CpSolverSolutionCallback):
    def __init__(self, exam_vars, exam_names, dates, log_func):
        cp_model.CpSolverSolutionCallback.__init__(self)
        
        self.__exam_vars = exam_vars
        self.__exam_names = exam_names
        self.__dates = dates
        
        self.__solution_count = 1
        self.__logger = log_func

    def on_solution_callback(self):
        """Called on each new solution."""
        obj = self.ObjectiveValue()
        bound = self.BestObjectiveBound()
        self.__logger(f'Feasible solution #{self.__solution_count} found, objective value = {obj}, best bound = {bound}')
        self.__solution_count += 1

        # save solution locally
        solution = extract_solution_from_solver(self, self.__exam_vars, self.__exam_names, self.__dates)
        write_solution_to_csv('schedule.csv', solution)

    def solution_count(self):
        """Returns the number of solutions found."""
        return self.__solution_count


### Read args
parser = argparse.ArgumentParser(description='TAU exam scheduler')
parser.add_argument('--secrets', 
                    help='TOML secrets file',
                    required=True)
parser.add_argument('--params', 
                    help='TOML params file',
                    required=True)
parser.add_argument('--debug', 
                    action='store_true',
                    default=False,
                    help='Printout solver log')
args = parser.parse_args()

# read config TOML files
with open(args.secrets, 'rb') as f:
    secrets = tomllib.load(f)
with open(args.params, 'rb') as f:
    params = tomllib.load(f)
debug = args.debug

# parameters
time_limit_in_mins = params['time_limit_in_mins']
absolute_gap_limit = params['absolute_gap_limit']
warm_start_prob = params['warm_start_prob']
dump_stats = params['log_stats']
dump_duplicates = params['log_duplicates']


#### Authorize and connect to Sheets ####
credentials = service_account.Credentials.from_service_account_info(
    secrets["gcp_service_account"],
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
    ],
)
gc = gspread.authorize(credentials)


#### Read Google Sheets input ####
sheet_url = secrets["private_gsheets_url"]
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

# Extract minimal and ideal gap constraints
sheet_name = 'מרווחים'
worksheet = workbook.worksheet(sheet_name)
data_rows = worksheet.get_all_values()[2:]

min_days_between_exams = {}
ideal_days_between_exams = {}
weights = {}
for row_i, row in enumerate(data_rows):
    pattern1, pattern2, min_days, ideal_days, weight = row[1].strip(), row[2].strip(), row[3].strip(), row[4].strip(), row[5].strip()
    if not (pattern1 and pattern2): continue

    pattern1 = preprocess_pattern(pattern1)
    pattern2 = preprocess_pattern(pattern2)
    pairs = get_matching_pairs(pattern1,pattern2,exam_names,exam_index)
    if len(pairs) == 0:
        log(f'Constraint in sheet {sheet_name}, row {row_i+3} yielded 0 matches')

    min_days = int(min_days) if min_days else None
    ideal_days = int(ideal_days) if ideal_days else None
    weight = int(weight) if weight else 1

    duplicates_found = False
    overriding = False
    for (exam1,exam2) in pairs:
        # ensure that exam1 < exam2 to avoid duplicates
        if exam1 == exam2: continue
        if exam1 > exam2: (exam1, exam2) = (exam2, exam1)
        pair = (exam1,exam2)

        # detect duplicates/overrides
        if pair in min_days_between_exams:
            duplicates_found = True
            if min_days and min_days != min_days_between_exams[pair]:
                overriding = True
            min_days_between_exams.pop(pair, None)
        
        if pair in ideal_days_between_exams:
            duplicates_found = True
            if min_days and min_days != ideal_days_between_exams[pair]:
                overriding = True
            ideal_days_between_exams.pop(pair, None)
            weights.pop(pair, None)

        # update values
        if min_days:
            min_days_between_exams[pair] = min_days
        if ideal_days:
            ideal_days_between_exams[pair] = ideal_days
            weights[pair] = weight
    
    if dump_duplicates and duplicates_found:
        if overriding:
            log(f'Duplicate constraint(s) detected in {sheet_name}, row {row_i+3} (OVERRIDING)')
        else:
            log(f'Duplicate constraint(s) detected in {sheet_name}, row {row_i+3} (non-overriding)')

# Filter out redundant constraints
for (pair, min_days) in min_days_between_exams.items():
    ideal_days = ideal_days_between_exams.get(pair)
    if ideal_days and ideal_days <= min_days: 
        # disable constraint
        ideal_days_between_exams[pair] = 0

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
        log(f'Constraint in sheet {sheet_name}, row {row_i+3} yielded 0 matches')

    duplicates_found = False
    for (exam1, exam2) in pairs:
        # detect duplicates
        if (exam1, exam2) in exam_before_exam:
            duplicates_found = True
            exam_before_exam.remove((exam1, exam2))

        exam_before_exam.append((exam1, exam2))
    
    if dump_duplicates and duplicates_found:
        log(f'Duplicate constraint(s) detected in {sheet_name}, row {row_i+3}')

# exam_before_date = []
# for row_i, row in enumerate(data_rows):
#     exam, date = row[18], row[19]
#     if exam and date:
#         exam = exam_index[exam]
#         date = date_index[date]
#         exam_before_date.append((exam, date))

# Extract prescheduled constraints
sheet_name = 'קיבועים'
worksheet = workbook.worksheet(sheet_name)
data_rows = worksheet.get_all_values()[2:]

exam_on_date = {}
for row_i, row in enumerate(data_rows):
    pattern, date = row[1].strip(), row[2].strip()
    if not (pattern and date): continue
    
    date = date_index.get(date)
    if not date:
        log(f'Invalid date in {sheet_name}, row {row_i+3}')
        continue

    pattern = preprocess_pattern(pattern)
    matches = get_matching(pattern,exam_names,exam_index)
    if len(matches) == 0:
        log(f'Constraint in sheet {sheet_name}, row {row_i+3} yielded 0 matches')
        continue

    duplicates_found = False
    for exam in matches: 
        # detect duplicates
        if exam in exam_on_date: duplicates_found = True
        exam_on_date[exam] = date

    if dump_duplicates and duplicates_found:
        log(f'Duplicate constraint(s) detected in {sheet_name}, row {row_i+3}')

# Add hints to existing solution if warmstart requested
hints = {}
if warm_start_prob > 0:
    sheet_name = 'שיבוץ'
    worksheet = workbook.worksheet(sheet_name)
    data_rows = worksheet.get_all_values()[3:]
    for row_i, row in enumerate(data_rows):
        exam, date = row[1].strip(), row[2].strip()
        if not exam or not date or not (exam in exam_index) or not (date in date_index): continue

        exam_i = exam_index[exam]
        date_i = date_index[date]
        hints[exam_i] = date_i

#### Construct scheduling problem ####

# Define the number of exams and the number of days
num_exams = len(exam_names)
horizon = len(dates)

# Create a CP-SAT model
model = cp_model.CpModel()

# Create variables
# exams = [model.NewIntVar(0, horizon-1, f'exam_{i}') for i in range(num_exams)]
exams = [None] * num_exams
for (exam_i,date_i) in exam_on_date.items():
    exams[exam_i] = date_i
for date_i in range(num_exams):
    if exams[date_i] is not None: continue
    exams[date_i] = model.NewIntVar(0, horizon-1, f'exam_{date_i}')

# Create intervals for each (exam,days) pair
gap_intervals = {}
for (i, j), days in min_days_between_exams.items():
    # ignore disabled constraints
    if days < 1: continue
    gap_intervals.setdefault((i,days), model.NewFixedSizeIntervalVar(exams[i], days, f'mingap_{i,days}'))
    gap_intervals.setdefault((j,days), model.NewFixedSizeIntervalVar(exams[j], days, f'mingap_{j,days}'))

# Add minimal gap constraints
for (i, j), days in min_days_between_exams.items():
    # ignore disabled constraints
    if days < 1: continue
    # Interval for each exam
    # interval_i = model.NewFixedSizeIntervalVar(exams[i], days, f'mingap_{i,j}')
    # interval_j = model.NewFixedSizeIntervalVar(exams[j], days, f'mingap_{j,i}')
    interval_i = gap_intervals[(i,days)]
    interval_j = gap_intervals[(j,days)]
    model.AddNoOverlap([interval_i, interval_j])
    # model.Add(exams[i] + min_days <= exams[j] or exams[j] + min_days <= exams[i])

# Add ideal gap constraints
ideal_violations = {}
for (i, j), days in ideal_days_between_exams.items():
    # ignore disabled constraints
    if days < 1: continue

    ideal_violations.setdefault((i,j), model.NewBoolVar(f'violation_{i,j}'))

    # Dedicated optional intervals for each pair of exams
    interval_i = model.NewOptionalFixedSizeIntervalVar(exams[i], days, ideal_violations[(i,j)].Not(), f'idealgap_{i,j}')
    interval_j = model.NewOptionalFixedSizeIntervalVar(exams[j], days, ideal_violations[(i,j)].Not(), f'idealgap_{j,i}')
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
    model.Add(exams[i] <= exams[j])

# # Add prescheduling constraints
# for (i,t) in exam_on_date.items():
#     model.Add(exams[i] == t)




# Define the objective

# Minimize soft constraints violation
# model.Minimize( sum(ideal_violations.values()) )

# Minimize soft constraints weighted violation
keys = ideal_violations.keys()
expr = [ideal_violations[k] for k in keys]
coef = [weights[k] for k in keys]
model.Minimize(cp_model.LinearExpr.WeightedSum(expr,coef))

# # Define the objective: makespan
# makespan = model.NewIntVar(0, horizon, 'makespan')
# model.AddMaxEquality(makespan, exams)
# model.Minimize(makespan)


# Add hints if warmstart requested
if warm_start_prob > 0:
    for (exam_i,date_i) in hints.items():
        # include hints at random
        if not (exam_i in exam_on_date) and random.random() < warm_start_prob:
            model.AddHint(exams[exam_i], date_i)

    # for (pair,bool_var) in ideal_violations.items():
    #     exam_i, exam_j = pair
    #     if exam_i == 0 or exam_j == 0: continue

    #     date_i, date_j = hints[exam_i], hints[exam_j]
    #     gap = abs(date_i - date_j)
    #     # min_days = min_days_between_exams.get(pair,'')
    #     ideal_days = ideal_days_between_exams.get(pair,None)
    #     if ideal_days:
    #         model.AddHint(bool_var, gap <= ideal_days)


# Create a solver and solve the model
solver = cp_model.CpSolver()
# Set solver parameters
if time_limit_in_mins > 0:
    solver.parameters.max_time_in_seconds = time_limit_in_mins * 60.0
if absolute_gap_limit > 0:
    solver.parameters.absolute_gap_limit = absolute_gap_limit
if debug:
    solver.parameters.log_search_progress = True
    solver.log_callback = print

# Solve!
log(f'Solving scheduling problem (time_limit_in_mins={time_limit_in_mins}, absolute_gap_limit={absolute_gap_limit})...')

solution_callback = MySolutionCallback(exams, exam_names, dates, log)
status = solver.SolveWithSolutionCallback(model, solution_callback)

log(f'Solver finished in {solver.WallTime()} s')
# status = solver.Solve(model)

# determine success & status
success = (status in [cp_model.OPTIMAL, cp_model.FEASIBLE])
status_name = solver.StatusName(status)
log(f'Solver status: {status_name}')


#### Save solution to the Google Sheet ####

# Write log to 'log' sheet
log_sheet = workbook.worksheet('log')
start_row = 1
end_row = log_sheet.row_count
range_name = f'A{start_row}:A{end_row}'
log_sheet.batch_clear([range_name])

log_sheet.update(range_name=range_name,
                 values=[logger], 
                 major_dimension='COLUMNS',
                 value_input_option="USER_ENTERED")

if success:
    # extract solution
    solution = extract_solution_from_solver(solver,exams,exam_names,dates)
    failed_list = extract_violations_from_solver(solver, ideal_violations, exams, exam_names, ideal_days_between_exams)

    # Write/backup solution to local csv file
    write_solution_to_csv('schedule.csv', solution)

    # Write output to 'שיבוץ' worksheet
    output = workbook.worksheet('שיבוץ')
    write_solution_to_gsheet(output, solution, failed_list)


if success and dump_stats:
    # calculate all gaps
    all_pairs = sorted( set().union(min_days_between_exams.keys(), ideal_days_between_exams.keys()) )
    
    data = []
    for pair in all_pairs:
        exam1, exam2 = pair
        name1, name2 = exam_names[exam1], exam_names[exam2]
        date1, date2 = solution.get(name1), solution.get(name2)
        if date1 is None or date2 is None: continue
        
        # actual_gap = abs((date1-date2).days)
        min_days = min_days_between_exams.get(pair,'')
        ideal_days = ideal_days_between_exams.get(pair,'')
        ideal_days = '' if ideal_days == 0 else ideal_days

        # actual gap will be computed by the spreadsheet
        # data.append([name1,name2,min_days,ideal_days,actual_gap])
        data.append([name1,name2,min_days,ideal_days])

    # Write output to 'debug' worksheet
    debug_sheet = workbook.worksheet('stats')

    # Clear existing content starting from row start_row
    start_row = 3
    end_row = debug_sheet.row_count
    debug_sheet.batch_clear([f'B{start_row}:E{end_row}'])

    # Write data
    debug_sheet.update(range_name=f'B{start_row}:E{start_row+len(data)-1}',
                  values=data, 
                  value_input_option="USER_ENTERED")
    