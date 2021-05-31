# General Physics Experiment / CSV Handler
import csv
import openpyxl as xl

def str_float(x):
    try:
        x = float(x)
    except ValueError:
        pass
    return x

filepath = '1_data.csv'

L = list()
with open(filepath, 'r', newline='') as f:
    for row in csv.reader(f, delimiter=','):
        # ws.append([str_float(x) for x in row])
        L.append([str_float(x) for x in row])

Runs = dict()
for i in range(len(L[0])):
    if L[0][i] in Runs:
        Runs[L[0][i]].append(i)
    else:
        Runs[L[0][i]] = [i]

wb = xl.Workbook()

for run in Runs.keys():
    run_collection = [list(i) for i in zip(*[[row[i] for row in L] for i in Runs[run]])]
    data_col_num = len(Runs[run])
    run_collection = [row for row in run_collection if row != ['']*data_col_num][1:]
    Runs[run] = run_collection
    ws = wb.create_sheet(run)
    ws.title = run
    for row in run_collection:
        ws.append(row)

wb.remove(wb['Sheet'])
wb.save(filepath[:-3]+'xlsx')
