import json
from signal import valid_signals
import xlsxwriter
from curses import raw
import pandas as pd
from pyparsing import empty
from pytest import skip
import numpy as np
import os
from datetime import datetime
from openpyxl import load_workbook


def write_json_file(ouput_file_path, json_string):
    # Directly from dictionary
    with open(ouput_file_path, 'w') as outfile:
        json.dump(json_string, outfile)
    print('Write into json file successfully: ' + ouput_file_path)


def get_report(data):
    # Reset variables
    report = ''
    try:
        #(1) Get the total no. of (i) sessions, (ii) messages, (iii) Fallback messages, and (iv) Fallback % in terms of all weeks
        #Checking the total messages & sessions

        #(i) seesions
        sum_sessions_arr = []
        for i in range(0,25):
            sum_sessions = 0
            for j in range(0,7):
                sum_sessions += data[i]["period"][j].get("no_session")
            sum_sessions_arr.append(sum_sessions)
        print("Total no. of users(conversations):",sum(sum_sessions_arr))
        df1 = pd.DataFrame(sum_sessions_arr,columns=['No.'])

        # (ii) messages
        sum_messages_arr = []
        sum_messages_month_dic = {"11/2021":0,"12/2021":0,"01/2022":0,"02/2022":0,"03/2022":0,"04/2022":0,"05/2022":0}
        sum_messages_year_dic = {"2021":0,"2022":0}
        for i in range(0,25):
            sum_messages = 0
            for j in range(0,7):
                sum_messages += data[i]["period"][j].get("no_message")
                for k in sum_messages_month_dic:
                    if data[i]["period"][j].get("date")[5:7] == k[0:2]:
                        sum_messages_month_dic[k] += data[i]["period"][j].get("no_message")
                for g in sum_messages_year_dic:
                    if data[i]["period"][j].get("date")[0:4] == g:
                        sum_messages_year_dic[g] += data[i]["period"][j].get("no_message")
            sum_messages_arr.append(sum_messages)
        print("Total no. of messages:",sum(sum_messages_arr))
        df2 = pd.DataFrame(sum_messages_arr,columns=['No.'])
        # df3 = pd.DataFrame(list(sum_messages_month_dic.items()),columns=['month','No.'])
        # df4 = pd.DataFrame(list(sum_messages_year_dic.items()),columns=['year','No.'])

        #(iii) FALLBACK messages
        sum_fallback_arr = []
        for i in range(0,25):
            sum_fallback_arr.append(data[i]["total_no_fallback"])
        print("Total no. of fallback:",sum(sum_fallback_arr))
        df5 = pd.DataFrame(sum_fallback_arr,columns=['No.'])
        
        #(iv) Fallback %
        print("Fallback %:",sum(sum_fallback_arr)/sum(sum_messages_arr)*100)

        # (2a) Get the total no. of sessions in terms of channels and weeks
        sum_sc_arr = {"app":0,"website":0,"msp":0,"others": 0}
        for i in range(0,25):
            for j in sum_sc_arr:
                sum_sc_arr[j] += data[i]["channel_session"][j]
        print("Total no. of conversations per channel:",sum_sc_arr)
        df6 = pd.DataFrame(list(sum_sc_arr.items()),columns=['channels','No.'])
        

        # (2b) Get the total no. of messages in terms of channels and weeks
        sum_mc_arr = {"app":0,"website":0,"msp":0,"others": 0}
        for i in range(0,25):
            for j in sum_mc_arr:
                sum_mc_arr[j] += data[i]["channel_msg"][j]
        print("Total no. of messages per channel:",sum_mc_arr)
        df7 = pd.DataFrame(list(sum_mc_arr.items()),columns=['channels','No.'])

        #(4a) Get the total no. of messages in terms of intents among all weeks
        lis_intent_dic = []
        for i in range(0,25):
            emp_dic = {}
            for j in range(0,len(data[i]["intent"])):
                emp_dic[data[i]["intent"][j][0]] = data[i]["intent"][j][1]
            lis_intent_dic.append(emp_dic)

        # collect all FAQs
        intent_emp_dic = lis_intent_dic[0].copy()
        for p in range(1,25):
            intent_emp_dic.update(lis_intent_dic[p])
        intent_emp_dic.update({}.fromkeys(intent_emp_dic,0))
        print("Total no. of FAQs:",len(intent_emp_dic))

        # start counting no. of each FAQs repeated
        intent_dic = intent_emp_dic.copy()
        for i in intent_dic:
            for j in range(0,25):
                for k in lis_intent_dic[j]:
                    if i == k:
                        intent_dic[i] += lis_intent_dic[j][k]
        sort_intent = sorted(intent_dic.items(), key=lambda x: x[1], reverse=True)

        # export to excel
        dic = intent_emp_dic.copy()
        arr = []
        for i in range(0,25):
            dic = intent_emp_dic.copy()
            for j in dic:
                for k in lis_intent_dic[i]:
                    if j == k:
                        dic[j] += lis_intent_dic[i][k]
            arr.append(dic)
        df8 = pd.DataFrame(arr)

        # #checking
        # arr = []
        # for i in range(0,25):
        #     arr.append(lis_intent_dic[i]["FALLBACK"])
        # print(arr)
        # print(sum(arr))

        # (5a) Get the rating count in terms of 1 to 5 stars in terms of all weeks
        rating = {"1":0, "2":0, "3":0, "4":0, "5":0}
        for i in range(0,25):
            for j in rating.keys():
                rating[j] = data[i]["rating"].get(j)+rating.get(j)
        print("Rating:",rating)

        # array of rating
        rating_arr = []
        m = {"1":0, "2":0, "3":0, "4":0, "5":0}
        for i in range(0,25):
            emp = m.copy()
            for j in m.keys():
                emp[j] = data[i]["rating"].get(j)
            rating_arr.append(emp)
        df9 = pd.DataFrame(rating_arr)

        # (6a) Usage distribution by hour
        hourly_usage = {"00:00-01:00":0,"02:00-02:00":0,"02:00-03:00":0,"03:00-04:00":0,"04:00-05:00":0,"05:00-06:00":0,
                        "06:00-07:00":0,"07:00-08:00":0,"08:00-09:00":0,"09:00-10:00":0,"10:00-11:00":0,"11:00-12:00":0,
                        "12:00-13:00":0,"13:00-14:00":0,"14:00-15:00":0,"15:00-16:00":0,"16:00-17:00":0,"17:00-18:00":0,
                        "18:00-19:00":0,"19:00-20:00":0,"20:00-21:00":0,"21:00-22:00":0,"22:00-23:00":0,"23:00-00:00":0}
        for i in range(0,25):
            k = 0
            for j in hourly_usage.keys():
                hourly_usage[j] = data[i]["msg_dist_hour_week"][k]+hourly_usage.get(j)
                k += 1
        df10 = pd.DataFrame(list(hourly_usage.items()),columns=['hour','No.'])
        print("Hourly usage:",hourly_usage)

        # (7a) Active & inactive seesions
        active_arr = []
        inactive_arr = []
        for i in range(0,25):
            inactive_arr.append(data[i]["no_active_session"])
            active_arr.append(sum_sessions_arr[i]-data[i]["no_active_session"])

        matrix_aux = np.vstack([active_arr,inactive_arr])
        matrix     = np.transpose(matrix_aux)
        df11 = pd.DataFrame(matrix,columns=['active', 'inactive'])
        print("Total no. of active conversations:",sum(active_arr))
        print("Total no. of inactive conversations:",sum(inactive_arr))

        # write all dataframe to one Excel
        writer = pd.ExcelWriter("weekly_data.xlsx", engine='xlsxwriter')
        df1.to_excel(writer,sheet_name='converations_weekly',index=False)
        df2.to_excel(writer,sheet_name='messages_weekly',index=False)
        df5.to_excel(writer,sheet_name='Fallback messages',index=False)
        df11.to_excel(writer,sheet_name='active&inactive_conv._weekly',index=False)
        df10.to_excel(writer,sheet_name='hourly_usage',index=False)
        df8.to_excel(writer,sheet_name='FAQs',index=False)
        df9.to_excel(writer,sheet_name='rating',index=False)
        df7.to_excel(writer,sheet_name='messages per channels_weekly',index=False)
        df6.to_excel(writer,sheet_name='conv. per channels_weekly',index=False)


        writer.save()

    except Exception as e:
        print("Error occurred")
        print(e)

    return report


def main():
    output_file_path = './report.json'

    data = []
    # (1) Loop over list of files to append to empty dataframe:
    with open('BD Chatbot Summary Week 1-26.json') as f1:
        summary_json = json.load(f1)

    # (2) Generate report
    report = get_report(summary_json)

    # (3) Write into json file
    write_json_file(output_file_path, report)


main()