import pandas as pd
import matplotlib.pyplot as plt
#%'matplotlib inline'
import datetime as dt
import pyodbc
import numpy as np
import win32com.client

# connect to SQL
db = 'Quality'
connection = pyodbc.connect('Driver={SQL Server}; Server=10.196.163.222,51433; Database=' + db + '; uid=; pwd=')
#cursor = connection.cursor()

# get current time for sql where statement
cur_time = dt.datetime.now()

# use previous month for analysis
# cur_time = dt.datetime.now() - dt.timedelta(days=30)

# sql command to get data
sql_command = """
SELECT  [Quality].[dbo].[tbl_ScrapData].[_id],
        [Quality].[dbo].[tbl_ScrapData].[Plnt],
        [Quality].[dbo].[tbl_ScrapData].[Year_month],
        [Quality].[dbo].[tbl_ScrapData].[PstngDate],
        [Quality].[dbo].[tbl_ScrapData].[Amtinloccur],
        [Quality].[dbo].[tbl_ScrapData].[AmountinDC],
        [Quality].[dbo].[tbl_ScrapData].[Material],
        [Quality].[dbo].[tbl_ScrapData].[Quantity],
        [Quality].[dbo].[tbl_ScrapData].[Text],
        [Quality].[dbo].[tbl_ScrapData].[AccountDescription],
        [Materials].[dbo].[tbl_MaterialList].[Material Desc]
        
  FROM [Quality].[dbo].[tbl_ScrapData]
  LEFT JOIN [Materials].[dbo].[tbl_MaterialList] ON [Quality].[dbo].[tbl_ScrapData].[Material]=[Materials].[dbo].[tbl_MaterialList].[Material]
  WHERE "Plnt" = 'A201' AND "Year_month" = '{}'
  """.format(cur_time.strftime('%Y/%m'))

# create pandas dataframe
df = pd.read_sql(sql_command, connection, index_col='_id')

# rename and convert from object type to float type
df['LC'] = df.Amtinloccur.astype(float)
df['DC'] = df.AmountinDC.astype(float)
df['Quantity'] = df.Quantity.astype(float)

# save df to file
#df.to_excel('scrap_data.xlsx')

# to group data to chart
df_slice = df[['PstngDate', 'LC', 'Material', 'Material Desc']]

df_dates_grp = df_slice.groupby('PstngDate')
df_dates_grp.sum()

df_slice_sum = df_slice['LC'].sum()
'${:,.2f}'.format(df_slice_sum)

# plot the scrap costs by date
ax = df_dates_grp.sum().plot(kind='bar', fontsize=(14), legend=False)
ax.set_xlabel("")
plt.title('Scrap $USD by Posting Date', fontsize=(20))
plt.tight_layout()
plt.savefig("C:\\Data\\posting_date_bar.png")

# limit the column width to make output eaiser to read
pd.set_option('display.max_colwidth', 25)

# create pivot table with data to display
df_dt_mat_pvt = pd.pivot_table(df, index=['PstngDate', 'Material', 'Material Desc', 'DC'], values=['LC', 'Quantity'], aggfunc=[np.sum])

# file = open('C:\\Data\\scrap_analysis-' + cur_time.strftime('%Y%m%d%H%M') + '.txt', 'w')

# create pivot
gt_500_lt_neg1000 = df_dt_mat_pvt.query('DC > 250 or DC < -500')
gt_500_lt_neg1000

df_neg_chrg = pd.pivot_table(df, index=['PstngDate', 'Material' , 'Text', 'AccountDescription', 'DC'], values=['LC'], aggfunc=[np.sum])
df_neg_chrg = df_neg_chrg.query('DC < -500')
df_neg_chrg

# email portion
mail_item = 0x0
obj = win32com.client.Dispatch('Outlook.Application')
new_mail = obj.CreateItem(mail_item)

mail_to_list = (''
                #"other@slb.com;" # to add or subtract members
               )

new_mail.Subject = 'Weekly Scrap Analysis'
new_mail.Attachments.Add('C:\\Data\\posting_date_bar.png') # must attach first before can make inline with <img> tag

body =  """
        <html>
            <body>
                <h2>Scrap analysis for: """ + cur_time.strftime('%Y-%m-%d %H:%M') + """</h2>
                <h3>Total scrap MTD: """ + '${:,.2f}'.format(df_slice_sum) + """</h3>
                <p> </p>
                <img src='cid:posting_date_bar.png' height=288 width=432>""" + df_dates_grp.sum().to_html() + """
                <h3>Scrap charges of note (>$250, <$-500):</h3>""" + gt_500_lt_neg1000.to_html() + """
                <h3>Negative charges of note:</h3>""" + df_neg_chrg.to_html() + """
            </body>
        </html>
        """

new_mail.HTMLBody = body                                                           
new_mail.To = mail_to_list
#new_mail.Save()
#new_mail.Send()
new_mail.Display()

df['DC'].sort_values()

## edit to make more OO

#function graph()

#function sql()

#function plot_format()

#function email()
#def email_output():

