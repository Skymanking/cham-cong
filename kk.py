import pandas as pd
data = pd.read_excel("OT.xlsx")
print(data.values[0][0])