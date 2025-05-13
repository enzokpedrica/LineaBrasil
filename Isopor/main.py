import pandas as pd

file_txt = "enzo.TXT"
file_csv =  "teste.XLS"

df = pd.read_csv(file_txt, skiprows=23, delimiter=";")