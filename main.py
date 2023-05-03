import pandas
from pulp import *
import pandas as pd
from openpyxl import load_workbook

# Integer programming for assigning topics to students
# according to students' preferences
# and upper limits on max number of students working on the same topic

# Read input from excel and create data frame. 
# The input file contains a sheet "input" with the students' preferences. The table organizes students in rows and topics in columns.
# The table entries represent the students' preferences for the given topics. A higher value stands for a higher preference.
# In my courses, usually, students may allocate 10 points on at least three topics. 
df=pd.read_excel("SemSolver.xlsx", "input")
df_short=df.iloc[:,3:df.shape[1]]
#points=df_short.to_dict('index')

# define list of students and topics, define max number of students for one topic
students=[str(i) for i in range(1,1+df_short.shape[0])]
topics=[str(i) for i in range(1,1+df_short.shape[1])]
# max number of students for one topic. Please try different values for this parameter based on number of students and number of topics in your course.
max_topics = 3

# Setting up optimization problem, decision variables
model = LpProblem("Maximize_group's_benefit", LpMaximize)
choices = LpVariable.dicts("choice",(students,topics),0,1,LpBinary)

# dictionary for budget points/ students' priorities
points = {}
for i in range(1,1+df_short.shape[0]):
    dict1={}
    for j in range(1,1+df_short.shape[1]):
        dict1[str(j)] = df_short.iloc[i-1,j-1]
    points[str(i)] = dict1

# Objective function
model += lpSum([points[i][j]*choices[i][j] for i in students for j in topics])

# Constraints for rows and columns of decision matrix
# rows: one topic per student
for i in students:
    model += lpSum([choices[i][j] for j in topics]) == 1
# columns: for each topic:
# students working on the topic less or equal than a given number (stored in max_topics)
for j in topics:
    model += lpSum([choices[i][j] for i in students]) <= max_topics

# solve the optimization problem
status=model.solve()

# display optimization model *******************************************************************************
print('*'*100)
print(model)
print('*'*100)
# display the results **************************************************************************************

# variable values
for var in model.variables():
    if var.value() > 0:
        print(f'Variable name: {var.name}, Variable value: {var.value()}')
print('*'*100)

# constraint values
for name, con in model.constraints.items():
    print(f'constraint name:{name}, constraint value:{con.value()}')
print('*'*100)

# objective value
print(f'OBJECTIVE VALUE IS: {model.objective.value()}')
print(f'students: {len(students)}, avg. benefit per student:{model.objective.value()/len(students)}')
print("Result:", LpStatus[status])

# store results in excel sheet *****************************************************************************

# create data frame (rows: students, columns: topics) with values of decision variables
df_results = pd.DataFrame.from_dict(choices, orient='index')
for i in students:
    for j in topics:
        df_results.loc[i,j]=value(choices[i][j])

# open excel file and write results to sheet 'results'
book=load_workbook('SemSolver.xlsx')
writer=pandas.ExcelWriter('SemSolver.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

df_results.to_excel(writer, sheet_name='results')
writer.save()
writer.close()









