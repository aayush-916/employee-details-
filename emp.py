import pandas as pd

# Read Excel file into a DataFrame
df = pd.read_excel('./emp.xlsx')
# Employees who has worked for 7 consecutive days

# Group by "Employee Name" and "Position Status," then count occurrences of "Position ID"
gr = df.groupby(["Employee Name", "Position Status"])["Position ID"].count()

# Filter and print Employee Names with counts greater than or equal to 7
filtered_employees = gr[gr >= 7]

# Creating a DataFrame with "Employee Name" and "Position Status" only
result_df = filtered_employees.reset_index()[["Employee Name", "Position Status"]]
# Saving it to excel 
result_df.to_excel('Output_work_for_7_consecutive.xlsx', index=False)


gr1 = pd.DataFrame(df)
gr1['Time'] = pd.to_datetime(df["Time"],format="%m/%d/%Y %I:%M %p")
gr1['Time Out'] = pd.to_datetime(df["Time Out"],format="%m/%d/%Y %I:%M %p")
gr1['Work'] = (gr1['Time Out'] - gr1['Time']).dt.total_seconds() / 3600
emp2 = gr1[(gr1['Work']<10) & (gr1['Work']>1)]
emp2 = emp2[['Employee Name','Position Status','Work']]
emp2.to_excel('Output_work_less_than_10hour.xlsx',index=False)

emp3 = gr1[(gr1['Work']>14)]
emp3 = emp3[['Employee Name','Position Status','Work']]


emp3.to_excel('Output_work_more_than_14hour.xlsx',index=False)