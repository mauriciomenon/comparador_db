import pandas as pd

data1 = pd.DataFrame({"x1":["x", "y", "x", "y", "y", "x", "y"],  # Create first DataFrame
                     "x2":range(3, 10),
                     "x3":["a", "b", "c", "d", "e", "f", "g"],
                     "x4":range(22, 15, - 1)})
print(data1) 

data2 = data1[0:0]
print(data2)                                           # Print second pandas DataFrame

d = data1.iloc[[1]]

data2 = pd.concat([data2,data1.loc[[1]]],ignore_index = False)

print(data2)   