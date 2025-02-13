import pandas as pd
import pulp


def OR_model(inhouse, OMP, Question2 = False):
    Demand = {}
    Inhouse_Capacity = {}
    
    OMP1_Cost = {}
    OMP2_Cost = {}
    OMP3_Cost = {}
    OMP1_Capacity = {}
    OMP2_Capacity = {}
    OMP3_Capacity = {}
    
    for i,row in inhouse.iterrows():
        Demand[row['Time (t)']] = row['D(t)']
        Inhouse_Capacity[row['Time (t)']] = row['PC(t)']
        
    for i,row in OMP.iterrows():
        OMP1_Cost[row['Time (t)']] = row['CO1(t)']
        OMP2_Cost[row['Time (t)']] = row['CO2(t)']
        OMP3_Cost[row['Time (t)']] = row['CO3(t)']
        OMP1_Capacity[row['Time (t)']] = row['OC1(t)']
        OMP2_Capacity[row['Time (t)']] = row['OC2(t)']
        OMP3_Capacity[row['Time (t)']] = row['OC3(t)']
       
    TIME = list(inhouse['Time (t)'])
    
    x = pulp.LpVariable.dict('x', TIME, lowBound= 0)
    y1 = pulp.LpVariable.dict('y1', TIME, lowBound= 0)
    y2 = pulp.LpVariable.dict('y2', TIME, lowBound= 0)
    y3 = pulp.LpVariable.dict('y3', TIME, lowBound= 0)
    z1 = pulp.LpVariable.dict('z1', TIME, cat = "Binary", lowBound= 0, upBound = 1)
    z2 = pulp.LpVariable.dict('z2', TIME, cat = "Binary", lowBound= 0, upBound = 1)
    z3 = pulp.LpVariable.dict('z3', TIME, cat = "Binary", lowBound= 0, upBound = 1)
    if Question2:
        s = pulp.LpVariable.dict('s', [0] + TIME, lowBound= 0)
        s[0] = 0
    
    if Question2:
        prob  = pulp.LpProblem('question2', pulp.LpMinimize)
    else:
        prob  = pulp.LpProblem('question1', pulp.LpMinimize)
    
    # objective
    prob += pulp.lpSum(OMP1_Cost[t]*y1[t] + OMP2_Cost[t]*y2[t] + OMP3_Cost[t]*y3[t] for t in TIME)
    
    # constraint: Demand should be met
    for t in TIME:
        if Question2:
            prob += x[t] + y1[t] + y2[t] +y3[t] + s[t-1] - s[t] >= Demand[t]
        else:
            prob += x[t] + y1[t] + y2[t] + y3[t] >= Demand[t]
        
    # constrain:capacity
    for t in TIME:
        prob += x[t] <= Inhouse_Capacity[t]
        prob += y1[t] <= z1[t]*OMP1_Capacity[t]
        prob += y2[t] <= z2[t]*OMP1_Capacity[t]
        prob += y3[t] <= z3[t]*OMP1_Capacity[t]
        
    # constraint: 5 times wala
    prob += pulp.lpSum(z1[t] for t in TIME) <= 5
    prob += pulp.lpSum(z2[t] for t in TIME) <= 5
    prob += pulp.lpSum(z3[t] for t in TIME) <= 5
    
    for t in TIME:
        # constraint: at most 2 OMPs at a time
        prob += z1[t] + z2[t] + z3[t] <=2
        # constraint: 30% wala
        prob += y1[t] + y2[t] +y3[t] <= 0.3*(x[t] + y1[t] + y2[t] +y3[t])
        
            
    prob.solve()
    optimality = pulp.LpStatus[prob.status]
    if optimality == 'Optimal':
        optimal_value = pulp.value(prob.objective)
        if Question2:
            print("Objective value for Q2: ", optimal_value)
        else:
            print("Objective value for Q1: ", optimal_value)
        
        solution_list = []
        for t in TIME:
            if Question2:
                this_time = {
                         'Time (t)': t,
                         'D(t)' : Demand[t],
                         'PC(t)' : Inhouse_Capacity[t],
                         'Inhouse_produced (t)': x[t].varValue,
                         'Outsourced_OMP1 (t)': y1[t].varValue,
                         'Outsourced_OMP2 (t)': y2[t].varValue,
                         'Outsourced_OMP3 (t)': y3[t].varValue,
                         'Stock Remains (t)':s[t].varValue
                        }
            else:
                this_time = {
                             'Time (t)': t,
                             'D(t)' : Demand[t],
                             'PC(t)' : Inhouse_Capacity[t],
                             'Inhouse_produced (t)': x[t].varValue,
                             'Outsourced_OMP1 (t)': y1[t].varValue,
                             'Outsourced_OMP2 (t)': y2[t].varValue,
                             'Outsourced_OMP3 (t)': y3[t].varValue,
                            }
            solution_list.append(this_time)
        solution = pd.DataFrame(solution_list)
        
        return optimality, solution
    else:
        return optimality, None
        
        
def main(filepath, output_path):
    file_path = pd.ExcelFile(filepath)
    Inhouse_df = pd.read_excel(file_path, sheet_name='Inhouse')
    OMP_df = pd.read_excel(file_path, sheet_name='OMP')
    soln_question1 = OR_model(Inhouse_df,OMP_df, False)
    soln_question2 = OR_model(Inhouse_df,OMP_df, True)
    with pd.ExcelWriter(output_path) as writer:
        if soln_question1[0]:
            soln_question1[1].to_excel(writer,sheet_name = 'Q1_Soln',index = None)
        if soln_question2[0]:
            soln_question2[1].to_excel(writer, sheet_name = 'Q2_Soln', index = None)
    return
    


if __name__ == "__main__":
    input_file_path = 'Satish Food Private Limited.xlsx'
    output_file_path = 'Assignment_Soln.xlsx'
    main(input_file_path,output_file_path)
