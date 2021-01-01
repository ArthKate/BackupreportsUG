
# coding: utf-8

# In[48]:


from PyPDF2 import PdfFileMerger
import matplotlib.pyplot as plt
from pandas.plotting import table
from datetime import date
from datetime import date, timedelta
import pandas as pd
import numpy as np
import os
import glob
# In[49]:


# *means all if need specific format then *.csv
list_of_files = glob.glob('D:\Backupreports\\UG\\*.CSV')
latest_file = max(list_of_files, key=os.path.getctime)
print(latest_file)
df = pd.read_csv(latest_file, skiprows=3)


# In[50]:


df['Job End Time'] = pd.to_datetime(df['Job End Time'], errors='coerce')

df["Job End Date"] = df['Job End Time'].dt.date


df["Job Start Time"] = pd.to_datetime(df['Job Start Time'], errors='coerce')

df["Job Start Date"] = df['Job Start Time'].dt.date


# In[51]:


df = df.dropna()

df["Status Code"] = df["Status Code"].astype(int)


df = df[df["Status Code"] != 150]


# In[52]:


m = pd.to_datetime('today') - timedelta(days=5)

x = pd.to_datetime('today')
x = str(x).split(" ")[0]
x1 = x+" 7:00:00"

y = pd.to_datetime('today') - timedelta(days=1)
y = str(y).split(" ")[0]
y1 = y+" 6:00:00"

k = pd.to_datetime('today') - timedelta(days=2)

k = k.strftime("%D %R")

m = m.strftime("%D %R")


# In[53]:


df["Job End Date"] = df["Job End Date"].astype(str)

df = df.replace(x, y)


# In[54]:


df = df[(df["Job End Time"].astype("datetime64[ns]") >= y1) &
        (df["Job End Time"].astype("datetime64[ns]") <= x1)]


# In[55]:


df["JOB_STATUS_COUNT"] = df.groupby(['Policy Name'])[
    'Job Status'].transform('nunique')


# In[56]:


df_final1 = df[(df["JOB_STATUS_COUNT"] >= 1) & (df["Job Status"] != "Failed")]


# In[57]:


df_final1 = pd.DataFrame(df_final1.groupby(["Policy Name", "Job Status"])[
                         ["Job End Date", "Job Primary ID"]].max().reset_index())


# In[58]:


df_final2 = df[(df["JOB_STATUS_COUNT"] == 1)]


# In[59]:


df_final3 = pd.DataFrame(df_final2.groupby(["Policy Name", "Job Status"])[
                         ["Job End Date", "Job Primary ID"]].max().reset_index())


# In[60]:


df_final = df_final1.append(df_final3)


# In[61]:


df_final = df_final.reset_index(drop=True).drop_duplicates()


# In[62]:


df_dump = df_final.merge(df.filter(
    ["Job Primary ID", "Schedule Type", "Status Code", "Job End Time"]), on=["Job Primary ID"])


# In[63]:


df_dump.drop(columns="Job End Date", inplace=True)


# In[64]:


df_dump.drop(columns=["Job Primary ID"], inplace=True)


# In[65]:


writer = pd.ExcelWriter('UG+_Backup_Data_'+x+'_.xlsx')


# In[66]:


#fig = ax.get_figure()


today = date.today()

x = str(today)


# fig.savefig("C:\\Users\\r84120261\\Desktop\\SW\\"+x+"_"+"SW.pdf",bbox_inches="tight")

df_dump1 = df_dump[(df_dump["Job Status"] == "Failed")]


# In[67]:


df_dump1


# In[68]:


# * means all if need specific format then *.csv
list_of_files = glob.glob('D:\Backupreports\\*.xlsx')
latest_file = max(list_of_files, key=os.path.getctime)
print(latest_file)
df_error = pd.read_excel(latest_file)


# In[69]:


df_dump1 = df_dump1.merge(df_error, on=["Status Code"], how="left")


# In[70]:


df_dump1


# In[71]:


df_dump1 = df_dump1.reset_index(drop=True).drop_duplicates()


# In[72]:cd /


# list_of_files = glob.glob('C:\\Users\\r84120261\\Desktop\\*.xlsx') # * means all if need specific format then *.csv
#latest_file = max(list_of_files, key=os.path.getctime)
#print (latest_file)
#df_inc= pd.read_excel(latest_file)


# In[73]:


# df_inc


# In[74]:


df_dump1.to_excel(writer, index=False, sheet_name="Failed_Policies")


# In[75]:


df_dump2 = df_dump[(df_dump["Job Status"] == "Successful") | (
    df_dump["Job Status"] == "Partially Successful")]


# In[76]:


df_dump2 = df_dump2.reset_index(drop=True).drop_duplicates()


# In[77]:


df_dump2.drop(columns="Status Code", inplace=True)


# In[78]:


df_dump2.to_excel(writer, index=False, sheet_name="Successful_Policies")


# In[79]:


writer.save()


# In[80]:


df_test3 = df_final.groupby(["Job End Date", "Job Status"])[
    "Job Status"].count().unstack("Job Status").fillna(0)


# In[81]:


df_test3


# In[82]:


s = set(df_test3.columns)
if "Partially Successful" not in s:
    df_test3["Partially Successful"] = 0


# In[83]:


if "Failed" not in s:
    df_test3["Failed"] = 0


# In[84]:


if "Successful" not in s:
    df_test3["Successful"] = 0


# In[85]:


df_test3["Successful"] = df_test3["Successful"] + \
    df_test3["Partially Successful"]


# In[86]:


df_test3.drop(columns="Partially Successful", axis=1, inplace=True)


# In[87]:


df_test3 = df_test3.tail(6)
df_test3 = df_test3.head(1)


# In[88]:


plt.subplots_adjust(left=0.2, top=0.8)


# In[89]:


df_test3["Failed"] = df_test3["Failed"].astype(int)
df_test3["Successful"] = df_test3["Successful"].astype(int)
df_test4 = df_test3.reset_index()

df_test4["Total"] = df_test4["Failed"]+df_test4["Successful"]

df_test4 = df_test4.reset_index()
df_test4.drop(["index"], axis=1, inplace=True)
df_test4 = df_test4.set_index("Job End Date")


# In[90]:


df_test3


# In[91]:


# colors_list = ['#5cb85c','#d9534f']
#ax=df_test3.plot(kind='bar', stacked=False,figsize=(15,8),width = 0.8,color = colors_list,edgecolor=None)
# ax.set_alpha(0.8)

# set individual bar lables using above list
# for i in ax.patches:
# get_x pulls left or right; get_height pushes up or down
# ax.text(i.get_x()+.04, (i.get_height())+8, \
# i.get_height().round(0), fontsize=11, color='black',
# rotation=45)
#ax.set_xlabel('Job Date', fontsize=14)
#ax.set_ylabel('Count', fontsize=15)


# In[92]:


colors_list = ['#d9534f', '#5cb85c']

fig, ax = plt.subplots(1, 1)
table = table(ax, np.round(df_test4, 2),
              bbox=(1.1, 0.4, 0.28, 0.28), colWidths=[0.3, 0.3, 0.3])
table.set_fontsize(22)
ax = df_test3.div(df_test3.sum(1), axis=0).plot(ax=ax, kind='bar', stacked=False,
                                                figsize=(15, 8), width=0.09, color=colors_list, edgecolor=None)
for p in ax.patches:
    width, height = p.get_width(), p.get_height()
    z, y = p.get_xy()
    handles, labels = ax.get_legend_handles_labels()
    lgd = ax.legend(handles, labels, loc='upper center',
                    bbox_to_anchor=(1.1, 0.9), fontsize=17)

    ax.annotate('{:.0%}'.format(height), (p.get_x()+.5*width,
                                          p.get_y() + height + 0.01), ha='center')
    ax.set_title("Uganda D2T BACKUP REPORT", fontsize=18)
    ax.set_xlabel('Job Start Date', fontsize=14)
    ax.set_ylabel('Count Percentage', fontsize=15)
    ax.tick_params(axis='both', which='major', labelsize=12.5)
    ax.tick_params(axis='both', which='minor', labelsize=12.5)


# In[93]:


fig.savefig("D2T_BACKUP_REPORT_"+x+"_UG.pdf", bbox_inches="tight")


# In[94]:


pdfs = ["D2T_BACKUP_REPORT_"+x+"_UG.pdf", "D2T_BACKUP_REPORT_"+x+"_SW.pdf",
        "D2T_BACKUP_REPORT_"+x+"_ZM.pdf", "D2T_BACKUP_REPORT_"+x+"_SS.pdf"]

merger = PdfFileMerger()

for pdf in pdfs:
    merger.append(pdf)

merger.write(x+"_"+"Backup_Report_D2T.pdf")
merger.close()


# In[ ]:


# html=df_final.to_html()

#import win32com.client
#outlook = win32com.client.Dispatch('outlook.application')
#mail = outlook.Createitem(0)
#mail.To = 'rakesh.kulkarni3@huawei.com'
#mail.Subject = 'auto mail python'
#mail.HTMLBody ="<html><body>Test image <img src=""C:\\Users\\r84120261\\Desktop\\test.png""></body></html>"

# mail.Send()
