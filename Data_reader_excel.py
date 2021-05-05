# -*- coding: utf-8 -*-
"""
Created on Fri Apr 30 15:08:56 2021

@author: Leo
"""

import serial as sr
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import Workbook
import xlsxwriter 
quest = int(input('Lasso di  tempo da monitorare:  (sec)  \n'))

massa= float(input('Inserisci peso: \n'))

file_save=input('File excel da salvare:  \n')

lasso= quest*10

s = sr.Serial('COM3', 9600);

plt.close('all');
plt.figure();
plt.ion();
plt.show();

data = np.array([]);



i = 0

index=2

dt=0.1


ass=[]
pot_list=[]
vel=[]
dt_list=[]

vel_v=0


while i < lasso :
    a = s.readline()
    c=a.decode()
    d = c.replace('aX =','')
    b_mill = int(d[0:6]);
    if b_mill > -280 and b_mill < 280 :
        b_mill=0
    b= (b_mill*9.81)/18851
    v=str(index) # trasformo index in una stringa per concatenarlo alla colonna in cui voglio scrivere il dato
    vel_v=b*0.1+vel_v
    forza = massa *b
    pot= forza * vel_v
    
    index +=1
    dt +=0.1
    
    #Collecting all data in lists
    
    ass.append(b)
    pot_list.append(pot)
    vel.append(vel_v)
    dt_list.append(dt)
    
    
    print(b)
    i +=1
    continue


file_name = file_save+'.xlsx'

workbook = xlsxwriter.Workbook(file_name)
worksheet1= workbook.add_worksheet()
worksheet2= workbook.add_worksheet()


worksheet1.write("A1","Accelerazione")
worksheet1.write("B1","Potenza")
worksheet1.write("C1","Velocità")
worksheet1.write_column("A2",ass)
worksheet1.write_column("B2",pot_list)
worksheet1.write_column("C2",vel)
worksheet1.write("D1","Delta Time")
worksheet1.write_column("D2",dt_list)

chart_ass = workbook.add_chart({"type": "line"})
chart_pot_list = workbook.add_chart({"type": "line"})
chart_vel = workbook.add_chart({"type": "line"})

lasso=str(lasso)
range_ass = "=Sheet1!$A$2:$A$"+lasso
range_pot_list = "=Sheet1!$B$2:$B$"+lasso
range_vel = "=Sheet1!$C$2:$C$"+lasso
range_dt = "=Sheet1!$D$2:$D$"+lasso

chart_ass.add_series({"values" : range_ass, "name": "aX","colour": "red"})
chart_pot_list.add_series({"values" : range_pot_list, "name": "|aX|","colour": "blue"})
chart_vel.add_series({"values" : range_vel, "name": "vX", "colour": "green"})

chart_ass.set_x_axis({"values":range_dt,'num_font':  {'name': 'Arial','size':30}})
chart_pot_list.set_x_axis({"values":range_dt,'num_font':  {'name': 'Arial','size':30}})
chart_vel.set_x_axis({"values":range_dt,'num_font':  {'name': 'Arial','size':30}})

worksheet2.insert_chart("A5", chart_ass, {'x_scale': 10, 'y_scale': 3})
worksheet2.insert_chart("A50", chart_pot_list,{'x_scale': 10, 'y_scale': 3})
worksheet2.insert_chart("A95", chart_vel,{'x_scale': 10, 'y_scale': 3})

workbook.close()

s.close()
