import os
import pandas as pd
from entsoe import EntsoePandasClient
from k import MY_KEY

# Get the present working directory
current_directory = os.path.dirname(os.path.abspath(__file__))

client = EntsoePandasClient(api_key=MY_KEY)  # get your API key at https://developer.entsoe.eu/

baltictimezone = 'Europe/Tallinn'
polishtimezone = 'Europe/Warsaw'

# Define the countries for which data will be fetched

countries = ['EE', 'LV', 'LT', 'FI', 'PL', 'DE', 'SE', 'NO', 'DK']
c1 = ['PL', 'DE', 'SE', 'NO', 'DK']


loaddata_file_path = os.path.join(current_directory, 'data/eledata_load_.xlsx')
productiondata_file_path = os.path.join(current_directory, 'data/eledata_production_.xlsx')

with pd.ExcelWriter(loaddata_file_path) as writer:
  for country in countries:
    if country in c1:
      t_zone = polishtimezone
    else:
      t_zone = baltictimezone

    start = pd.Timestamp('20180101', tz=t_zone)
    end = pd.Timestamp('20250228', tz=t_zone)

    data = client.query_load(country, start=start, end=end)

    data_1hr = data.resample('h').mean().tz_localize(None)
    yearly_data = data_1hr.resample('YE').sum()
    monthly_data = data_1hr.resample('ME').sum()
    yearly_data.to_excel(writer, sheet_name=country+'_yearly')
    monthly_data.to_excel(writer, sheet_name=country+'_monthly')
    print('Country: ', country)
    print(yearly_data)
    print('------------------------------------') 
    

with pd.ExcelWriter(productiondata_file_path) as writer:
  for country in countries:
    if country in c1:
      t_zone = polishtimezone
    else:
      t_zone = baltictimezone

    start = pd.Timestamp('20180101', tz=t_zone)
    end = pd.Timestamp('20250228', tz=t_zone)

    data = client.query_generation(country, start=start, end=end, psr_type=None)

    data_1hr = data.resample('h').mean().tz_localize(None)
    yearly_data = data_1hr.resample('YE').sum()
    monthly_data = data_1hr.resample('ME').sum()
  
    # Write the yearly data to a sheet named after the country
    yearly_data.to_excel(writer, sheet_name=country+'_yearly')
    monthly_data.to_excel(writer, sheet_name=country+'_monthly')
        
    print('Country: ', country)
    print(yearly_data)
    print('------------------------------------')

