import pandas as pd
from vba_password import set_password
from datetime import date

# Fix for running Excel in "non-interactive mode", you require two directories to be created.
# You have to create a folder (or two on a 64bit-windows):
# (32Bit, always) C:\Windows\System32\config\systemprofile\Desktop
# (64Bit) C:\Windows\SysWOW64\config\systemprofile\Desktop
# If these directories are present, MS Excel will open in interactive mode rather than non-interactive.
# This then allows scheduled tasks to run flawlessly.

today = date.today()

df = pd.read_csv(r"C:\Users\Gareth.Cadman\Desktop\data.csv",
                 header=0)

df = df.sort_values(by='Date', ascending=False)

SaveAs = rf"C:\Users\Gareth.Cadman\Desktop\data {today}.xlsx"

writer = pd.ExcelWriter(SaveAs)
df.to_excel(writer, sheet_name='pandas with pass', index=False)
writer.save()
writer.close()

# Use imported set_password function to put a password on the SaveAs file.
set_password(SaveAs, 'pandas')
