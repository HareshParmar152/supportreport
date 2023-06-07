import streamlit as st
import pandas as pd
from datetime import datetime

# File upload
st.title('Support Report')
file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if file is not None:
    try:
        # Read Excel file
        df = pd.read_excel(file)
        #Agent Name
        Name = df['Agent Name'][1]
        #Report Date
        temp = df['Start Time'].iloc[-1]
        Date = temp.split(' ')
        Date = Date[0]
        date_format = "%Y-%m-%d"
        temp = datetime.strptime(Date, date_format)

        # Current date and time
        now = temp

        # Convert to Indian date format
        indian_date_format = now.strftime("%d-%m-%Y")


        #First call
        temp = df['Start Time'].iloc[-1]
        f_time = temp.split(' ')
        f_time = f_time[1]

        from datetime import datetime

        # Assuming you have a string representing the time in 24-hour format
        time_24h = f_time

        # Convert the time to AM/PM format with hours, minutes, and seconds
        f_time = datetime.strptime(time_24h, "%H:%M:%S").strftime("%I:%M:%S %p")

        #Last Call time
        temp = df['End Time'].iloc[1]
        l_time = temp.split(' ')
        l_time = l_time[1]

        from datetime import datetime

        # Assuming you have a string representing the time in 24-hour format
        time_24h = l_time

        # Convert the time to AM/PM format with hours, minutes, and seconds
        l_time = datetime.strptime(time_24h, "%H:%M:%S").strftime("%I:%M:%S %p")

        T_Inbound_calls = len(df[(df['Disposition']=='ANSWER') & (df['Call Type'] == 'INBOUND')])

        T_Outbound_calls = len(df[(df['Disposition']=='ANSWER') & (df['Call Type'] == 'OUTBOUND')])

        T_Answer_calls = len(df[df.Disposition == 'ANSWER'])

        T_Duration = round(sum(df['Duration (in seconds)'])/60,2)
        
        # Display DataFrame
        st.dataframe("Name : {}\nDate : {}\nShift time :\nPreference : \nLive : \nOffline : \n\nFirst call : {}\nLast call : {}\nStart Break Time : \nEnd Break Time : \nStart Short Break Time : \nEnd Short Break Time : \nInbound calls : {}\nOutbound calls : {}\nTotal calls : {}\nProductive Minutes : {}"
      .format(Name,indian_date_format,f_time,l_time,T_Inbound_calls,T_Outbound_calls,T_Answer_calls,T_Duration))
    except Exception as e:
        #st.error("Error: Failed to read the Excel file.")
        st.error(str(e))
