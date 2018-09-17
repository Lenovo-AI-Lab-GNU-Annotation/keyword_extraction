#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Aug 14 09:51:22 2018

@author: chenjian
"""

import re
import csv
import sys
import os

if len(sys.argv)>1:
    arr = os.listdir(sys.argv[1])
else:
    arr = os.listdir()

#print (arr)

arr_csv = [b for b in arr if b[-4:]=='.csv']

print (arr_csv)

#%%

def get_list_of_vib_sents(csv_file_name):
    
    csv_file = open(csv_file_name,encoding='ISO-8859-1')
    
    ws_input = csv.reader(csv_file)
    
    data = [r for r in ws_input]
    
    csv_file.close()
    
    vib_logs = []

    list_logs = []
    
    for i in data[2:]:
        
        chatid = i[0]
        
        chatlog = i[16]
        
        #print(chatid, chatlog)
        
        first_line = re.split('\n',chatlog)[0]
        
        pos_0 = first_line.find("my name is ")+11
        
        name_agent = first_line[pos_0:pos_0+4]
        
        if name_agent == "":
            print (chatlog)
            continue
            
      #  print (name_agent)
        
        chatturns = re.split('\n',chatlog)
    
        for index,chatturn in enumerate(chatturns):
            
            #print (chatturns)
            
            if name_agent in chatturn:
                continue
            
            if 'charger' in chatturn:
                #print (name_agent,chatturn)
                
                vib_logs.append(chatturn)
                
                item = chatturn
                
                list_words = item.split()
                
                flag_skip = 0
                
                for n_word in ['replace','replacement','buy','bought','you','your','we','wireless','different','other','multiple']:
                    
                    if n_word in item.split():
                        
                        flag_skip = 1
            
                
                if flag_skip ==1:
                    continue
                
                sent = item[item.find(':',10)+2:]

                if flag_skip != 1:
                    list_logs.append((str(chatid)+"-"+str(index),sent,len(sent.split())))
                    
                continue

    
    return list_logs

#%%
    
wb_output = load_workbook(filename = '/Users/chenjian/Lenovo/20180704_preExtractCase_new/chatturn_temp.xlsx')
ws_output = wb_output.worksheets[0]

output_row = 2
#%%    
for item in arr_csv:

    output_list = set(get_list_of_vib_sents(item))
    
    for tuple_output in output_list:
        
        if tuple_output[2]<5:
            continue
        
        ws_output['A'+str(output_row)].value = tuple_output[0]
        ws_output['B'+str(output_row)].value = tuple_output[1]
        ws_output['C'+str(output_row)].value = tuple_output[2]
        
        print (output_row, tuple_output)
        
        output_row +=1

#%%
 
wb_output.save(filename = '/Users/chenjian/Lenovo/0626/charging_port/charger_csv.xlsx')
