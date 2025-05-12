import pandas as pd

file = 'isopor.TXT'

df = pd.read_csv(file, delimiter=';', encoding='latin1')
df1 = pd.DataFrame(df)
print(df1.columns)