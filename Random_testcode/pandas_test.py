import pandas as pd
import numpy as np
from configobj import ConfigObj
from datetime import datetime, date, timedelta
import os.path

config = ConfigObj('config.ini')
d1 = date(2017, 1, 1)  # start date
d2 = date(2017, 1, 31)  # end date
column = []
dates = []
delta = d2 - d1         # timedelta
date = datetime.today()
today = date.strftime('%d-%m-%Y')
test_today = '19-06-2017'

index_test = [1, 2, 3, 4, 5, 6, 7, 8, 9]
index_test2 = [1, 2]
values1 = [1, 0, 1, 0, 0, 0, 0, 0, 0]
values2 = [0, 1, 0, 1, 0, 0, 0, 0, 0]
values3 = [0, 0, 1, 0, 1, 0, 0, 0, 0]
values4 = [1, 0, 1, 1, 0, 1, 0, 0, 0]

test1 = np.array([1, 0, 1, 0, 0, 0, 0, 0, 0])

for i in config['Vedlikeholdspunkt'].values():
    column.append(i)


if os.path.isfile('pandas_to_excel.xlsx') == False:
    df = pd.DataFrame(values1, index=column, columns=[today])

else:
    df = pd.read_excel('pandas_to_excel.xlsx', sheet_name='Sheet1')


#Litterally just slapping the values on to the existing dataframe! Just how we like it!
df[test_today] = pd.Series(values3, index=df.index)

# new_data = pd.DataFrame(values2, index=column, columns=['Hello'])
#
# test = pd.merge(df, new_data, left_index=True, right_on='Hello')
# new_df = new_df.set_index(column, inplace=True)

#df.fillna(value=0)
# for i in range(delta.days + 1):
#     dates.append(d1 + timedelta(days=i))

#dates2 = pd.date_range('20130101', periods=6)

#df = pd.DataFrame(np.random.randn(6,4), index=dates, columns=list('ABCD'))
#df = pd.DataFrame(config['Vedlikeholdspunkt'])
#df = pd.DataFrame(np.random.randn(365,9),index=d1, columns=column)
#df2 = df['Vedlikeholdspunkt']
#df2.set_index([1, 2, 3, 4, 5, 6, 7, 8, 9])
#df.set_index()


    #print(d1 + timedelta(days=i))


#today = datetime.today()
#df.to_excel('pandas_to_excel.xlsx', sheet_name='Sheet1')
#test.to_excel('pandas_to_excel.xlsx', sheet_name='Sheet1')

print(df)
