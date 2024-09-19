from datetime import datetime, timedelta
   
today = datetime.today()
str_today = str(today)
date_today = str_today.replace("-","")
new_date_today = date_today.split(' ',1)[0]
yyyymmdd = new_date_today.replace(" ","")
current_month = today.month
current_year = today.year

print(yyyymmdd)