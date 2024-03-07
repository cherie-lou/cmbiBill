import datetime as dt
import pandas as pd
import streamlit as st
import openpyxl
import xlsxwriter
import io

st.title("cmbi账单Read")
st.header('File Input')
input_file = st.file_uploader("Upload an Excel File", type=["xls", "xlsx"])

if input_file is not None:
    df = pd.read_excel("/Users/CherieLou/Downloads/cmbi账单.xlsx",sheet_name='本期订单')
    df = df[['乘车人姓名','企业实付金额','实际出发地','实际目的地','用车备注','补充说明','开始计费时间','结束计费时间','下单时间']]
    df['开始计费时间'] = pd.to_datetime(df['开始计费时间'])
    df['结束计费时间'] = pd.to_datetime(df['结束计费时间'])
    df['下单时间'] = pd.to_datetime(df['下单时间'])
    df['date'] = df['下单时间'].apply(dt.datetime.date)
    dic = {}
    grouped = df.groupby(['乘车人姓名','date'])
    for (name,time),s in grouped:
        time_str = dt.datetime.strftime(time,"%Y-%m-%d")
        for i,trip in s.iterrows():
            if trip['用车备注'] == "商务出行":
                continue
            elif trip['用车备注'] == "出差":
                try:
                    if '机场' in trip['实际出发地'] or '进站口' in trip['实际出发地'] or '出站口' in trip['实际出发地']:
                        continue
                    if '机场' in trip['实际目的地'] or '进站口' in trip['实际目的地'] or '出站口' in trip['实际目的地']:
                        continue
                    
                    dic[(name,time_str)] = dic.get((name,time_str),0) + trip['企业实付金额']
                except:
                    print(trip)
            else:
                dic[(name,time_str)] = dic.get((name,time_str),0) + trip['企业实付金额']  
    limit = 150
    over_limit = {}
    for pair,amount in dic.items():
        if amount > limit:
            over_limit[pair] = amount 
    temp = []
    for (name,date),amount in over_limit.items():
        temp.append([name,date,amount])
    result = pd.DataFrame(temp,columns=['姓名','日期','金额'])
    output_final = io.BytesIO()  # Create a bytes buffer to store the Excel file
    with pd.ExcelWriter(output_final, engine='xlsxwriter') as writer:
        result.to_excel(writer)
    excel_data_final = output_final.getvalue()
    st.download_button(label="Download Excel File", data=excel_data_final, file_name="result.xlsx", key='download')
