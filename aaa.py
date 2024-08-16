WORKTIME_TYPE_COL = 'C'
WORKTIME_START_DAY_COL = 'D'
WORKTIME_START_TIME_COL = 'E'
WORKTIME_END_DAY_COL = 'F'
WORKTIME_END_TIME_COL = 'G'
WORKTIME_STARTING_ROW = 10


import xlwings as xw

# Configurable data
TEMPLATE_NAME = 'LastName_FirstName_WorkTimeRecord_2024-08.xlsx'
FIRST_NAME = 'Suren'
LAST_NAME = 'Abeghyan'
GROUP = 'Undulator Systems'
TARGET_MONTH = 9
TARGET_YEAR = 2024

# Structure constants
PLAN_DAYTYPE_COL = 'D'
PLAN_DAYOFMONTH_COL = 'A'
PLAN_STARTING_ROW = 13
WORKTIME_TYPE_COL = 'C'
WORKTIME_START_DAY_COL = 'D'
WORKTIME_START_TIME_COL = 'E'
WORKTIME_END_DAY_COL = 'F'
WORKTIME_END_TIME_COL = 'G'
WORKTIME_STARTING_ROW = 10

# Open the workbook
workbook = xw.Book(TEMPLATE_NAME)

# Identify the worksheets
worksheet_profile = workbook.sheets['My Profile']
worksheet_plan = workbook.sheets['Monthly Plan and Absences']
worksheet_time = workbook.sheets['Enter Working Time']

# Profile
worksheet_profile.range('C3').value = f'{FIRST_NAME} {LAST_NAME}'
worksheet_profile.range('C4').value = GROUP

# Change the target month
worksheet_plan.range('C5').value = TARGET_MONTH

# Figure out start working day of month
working_days = []
for row in range(PLAN_STARTING_ROW, PLAN_STARTING_ROW+31):
    week_day = worksheet_plan.range(f'{PLAN_DAYTYPE_COL}{row}').value
    print(week_day)
    if week_day == 'Working day':
        working_days.append(int(worksheet_plan.range(f'{PLAN_DAYOFMONTH_COL}{row}').value))

if len(working_days) == 0:
    print('Unable to detect working days of the month')
    exit(1)

print(f'Working days in month: {working_days}')

# Start filling working times
worktime_row = WORKTIME_STARTING_ROW
for day in working_days:
    worksheet_time.range(f'{WORKTIME_TYPE_COL}{worktime_row}').value = 'Office Hours'
    worksheet_time.range(f'{WORKTIME_START_DAY_COL}{worktime_row}').value = day
    worksheet_time.range(f'{WORKTIME_START_TIME_COL}{worktime_row}').value = '08:00'
    worksheet_time.range(f'{WORKTIME_END_DAY_COL}{worktime_row}').value = day
    worksheet_time.range(f'{WORKTIME_END_TIME_COL}{worktime_row}').value = '12:30'

    worksheet_time.range(f'{WORKTIME_TYPE_COL}{worktime_row + 1}').value = 'Office Hours'
    worksheet_time.range(f'{WORKTIME_START_DAY_COL}{worktime_row + 1}').value = day
    worksheet_time.range(f'{WORKTIME_START_TIME_COL}{worktime_row + 1}').value = '13:00'
    worksheet_time.range(f'{WORKTIME_END_DAY_COL}{worktime_row + 1}').value = day
    worksheet_time.range(f'{WORKTIME_END_TIME_COL}{worktime_row + 1}').value = '15:30'

    worksheet_time.range(f'{WORKTIME_TYPE_COL}{worktime_row + 2}').value = 'Remote work'
    worksheet_time.range(f'{WORKTIME_START_DAY_COL}{worktime_row + 2}').value = day
    worksheet_time.range(f'{WORKTIME_START_TIME_COL}{worktime_row + 2}').value = '16:10'
    worksheet_time.range(f'{WORKTIME_END_DAY_COL}{worktime_row + 2}').value = day
    worksheet_time.range(f'{WORKTIME_END_TIME_COL}{worktime_row + 2}').value = '16:48'

    worktime_row += 3

# Save the workbook with a new name
workbook.save(f'{LAST_NAME}_{FIRST_NAME}_WorkTimeRecord_{TARGET_YEAR}-{TARGET_MONTH:02}.xlsx')

# Close the workbook
workbook.close()