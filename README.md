# BackupreportsUG

My repo folder I have the UG folder with the raw .csv file 
Root folder contains the script file
purpose is to generate and .xlsx and a pdf report
upon running the script file, I get response below


SyntaxError: invalid syntax
PS D:\Backupreports> python FINAL-BACKUP_REPORT_UG-SINGLE.py
D:\Backupreports\UG\OpsCenter_Backup-Automated-Report-Daily-_Uganda_01_01_2021_07_05_21_492_AM_78.CSV
D:\Backupreports\UG+_Backup_Data_2021-01-01_.xlsx
Traceback (most recent call last):
  File "FINAL-BACKUP_REPORT_UG-SINGLE.py", line 186, in <module>
    df_error = pd.read_excel(latest_file)
  File "C:\Users\kwx644959\AppData\Local\Programs\Python\Python38\lib\site-packages\pandas\util\_decorators.py", line 299, in wrapper
    return func(*args, **kwargs)
  File "C:\Users\kwx644959\AppData\Local\Programs\Python\Python38\lib\site-packages\pandas\io\excel\_base.py", line 336, in read_excel
    io = ExcelFile(io, storage_options=storage_options, engine=engine)
  File "C:\Users\kwx644959\AppData\Local\Programs\Python\Python38\lib\site-packages\pandas\io\excel\_base.py", line 1062, in __init__
    ext = inspect_excel_format(
  File "C:\Users\kwx644959\AppData\Local\Programs\Python\Python38\lib\site-packages\pandas\io\excel\_base.py", line 954, in inspect_excel_format    raise ValueError("File is not a recognized excel file")
ValueError: File is not a recognized excel file

I need assistance to figure out the output is not being recognized as an excel file. 
All project files attached
