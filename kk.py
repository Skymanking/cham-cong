import pandas as pd
data = pd.read_excel("data1.xlsx")
print(data.values[1][0])