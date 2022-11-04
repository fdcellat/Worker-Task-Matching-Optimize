# -*- coding: utf-8 -*-
"""
Created on Thu Sep 22 11:21:15 2022

@author: T006940
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Sep 21 15:24:39 2022

@author: T006940
"""

from ortools.linear_solver import pywraplp
import pandas as pd

column_names=['worker','task','time']
output=pd.DataFrame(columns=column_names)

df=pd.read_excel('dataopt.xlsx')
costs=df.values.tolist()
# costs = [
#        [14,5,8,7,3],
#        [3,2,1,7,4],
#        [2,12,6,5,2],
#        [7,8,3,9,1,]
#     ]

def main():
    # Data
    global output
    num_workers = len(costs)
    num_tasks = len(costs[0])

    # Solver
    # Create the mip solver with the SCIP backend.
    solver = pywraplp.Solver.CreateSolver('SCIP')

    if not solver:
        return

    # Variables
    # x[i, j] is an array of 0-1 variables, which will be 1
    # if worker i is assigned to task j.
    x = {}
    for i in range(num_workers):
        for j in range(num_tasks):
            x[i, j] = solver.IntVar(0, 1, '')

    # Constraints
    # Each worker is assigned to at most 2 min 1 task.
    for i in range(num_workers):
        solver.Add(solver.Sum([x[i, j] for j in range(num_tasks)]) <= 2)
    for i in range(num_workers):
        solver.Add(solver.Sum([x[i, j] for j in range(num_tasks)]) >= 1)

    # Each task is assigned to exactly one worker.
    for j in range(num_tasks):
        solver.Add(solver.Sum([x[i, j] for i in range(num_workers)]) == 1)
    

    # Objective
    objective_terms = []
    for i in range(num_workers):
        for j in range(num_tasks):
            objective_terms.append(costs[i][j] * x[i, j])
    solver.Minimize(solver.Sum(objective_terms))

    # Solve
    status = solver.Solve()

    # Print solution.
    if status == pywraplp.Solver.OPTIMAL or status == pywraplp.Solver.FEASIBLE:
        print(f'Total cost = {solver.Objective().Value()}\n')
        for i in range(num_workers):
            for j in range(num_tasks):
                # Test if x[i,j] is 1 (with tolerance for floating point arithmetic).
                if x[i, j].solution_value() > 0.5:
                    output=output.append({'Calisan_No':i,'Banka_No':j,'Süreler':costs[i][j]},ignore_index=True)
                    #print(f'Worker {i} assigned to task {j}.' +
                          #f' Cost: {costs[i][j]}')
    else:
        print('No solution found.')


if __name__ == '__main__':
    main()
    
output.to_excel('minimize.xlsx',index=False)