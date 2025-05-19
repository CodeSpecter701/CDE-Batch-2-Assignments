import pandas as pd

df = pd.DataFrame([["A", 1], ["B", 2]], columns=["Letter", "Number"])
df.to_excel("test_output.xlsx", index=False, engine="openpyxl")
print("Excel file created!")
