import gurobipy as gb

m = gb.Model("problem")

x = m.addVars(2, lb=0, ub=1, vtype=gb.GRB.BINARY, name="x")
y = m.addVars(1, lb=-5, ub=100, vtype=gb.GRB.INTEGER, name="y")

m.setObjective(x[0] + x[1]**2 + y[0], gb.GRB.MAXIMIZE)

m.optimize()
