# -*- coding: utf-8 -*-
"""
Spyder Editor
"""
import numpy as np
import itertools
import xlwt
from xlwt import Workbook
import math
from openpyxl import load_workbook

#name the workbook 
wb = Workbook()

#Add a sheet and name it. Allow cell overwrite
permwb = wb.add_sheet('gearpermutation', cell_overwrite_ok=True)

#initiating my variables

#Gear names - can be changed to whatever you'd like to label the gears
gearnames = ['compound type','Gear 1', 'Gear 2', 'Gear 3', 'Gear 4', 'Gear 5']

values = [2, 4, 6, 8, 10] #array of gears w desired teeth numbers

per = [] #storage of the number of permultations given len(values) choose i
per1 = [] #solution array
speed = ['speed'] #the speed ratio of the gears

#For loop which stores all the permutations in which you can choose i of the length of the values array
for choose in range(len(values)+1):
    #per is a forced list of a map which forces all the tuples in itertools to lists in python - a bit clunky, but functions
    per =  list(map(list, itertools.permutations(values, choose)))

    #checks if i > 1. i < 1 means that there are not two gears to compound
    if choose > 1: 
        #itereates over the number of required replications of possible compound gears 
        for rep in range(choose):
            #iterates over the length of the permutations given len(values) choose (choose) from above
            for index in range(len(per)):               
                per1.append(list(per[index])) #appends per at a given index to the solution array
                per1[-1].insert(0,rep) #inserts the replication number to track which combo is happening - not elegant 
                
                #Appends the calculated speed to the speed array, checks if rep = 0
                #checks if the replication is 0, because the calculation for no compunds is different than with compounds
                if rep ==0:
                    speed.append(per[index][rep]/per[index][-1])
                else:
                    speed.append(per[index][0]/per[index][rep-1] * per[index][rep]/per[index][-1])
                
    else: 
        #iterates over the length of per in the case that there are no possible compound gears to be made
        for length in range(len(per)): 
            speed.append(1) #adds the only speed reduction possible with less than 2 gears
            per1.append(per[length]) #appends a given permutation to the solution array
            per1[-1].insert(0,0) #adds the 'combination' type


print(len(per1), len(speed))
per1.insert(0, gearnames[0:len(gearnames)])
#print(per1)

#iterates over the 2d array of the output and prints it to an excel workbook. 
for i in range(len(per1)):
    permwb.write(i, len(gearnames), speed[i])
    for n in range(len(per1[i])):
        permwb.write(i,n, per1[i][n])

#saves the workbook        
wb.save('20210613_gearpermutation2.xls')
