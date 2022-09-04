import pandas as pd
import urllib
from sqlalchemy import create_engine

print('opening excel')

xlsx_path = (r"/home/king/Desktop/sample-entry.xlsx")
df = pd.read_excel(xlsx_path)

print(df.head(10))

print('opening access....')

cnnstr = (
    r"Driver = (Microsoft Access Driver (*.mdb, *.accdb))",
    r"DBQ = /home/king/Desktop/new.mdb"
)

cnnurl = f"access+pyodbc:///?odbc_connect={urllib.parse.quote(str(cnnstr))}"
acc_engine = create_engine(cnnurl)

print('writing to access... (new table)')

df.to_sql('new_access', con=acc_engine)

print('write comlpate')