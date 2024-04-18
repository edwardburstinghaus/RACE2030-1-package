## Import needed libraries
import powerfactory as pf
import importlib
import pandas as pd
import numpy as np
import datetime
from openpyxl import Workbook
import random

## Get all the contents of the script 
app = pf.GetApplication()
app.ClearOutputWindow()
# app.EchoOff()
script = app.GetCurrentScript()
Grid = script.Grid

## Construct array of all LV loads
all_LV_loads = app.GetCalcRelevantObjects('ElmLodlv')

## This variables defines the target for the average number of EVs per household in the network; note that this should always be set at a value above
## the current average number of EVs per household encoded into the model 
target_residential_avg_EV_count = 0.053136

## This variable defines the assumed maximum number of EVs that would be owned by any one household
upper_EV_limit_per_household = 2

## Total number of households
total_no_households = 13535

## Read in some data that is needed to construct the customers to phases and customers to loads maps
wb = pd.read_excel('LV_loads_to_PPNs_map.xlsx').values
customers_PPN_list=[]
A_N_load_list=[]
B_N_load_list=[]
C_N_load_list=[]
ABC_load_list=[]
A_B_load_list=[]
A_C_load_list=[]
B_C_load_list=[]
customers_to_loads_map=[]
customers_to_phases_map=[]
chosen_so_far=[]
residential=[]
for item in wb:
    customers_PPN_list.append(item[0])
    A_N_load_list.append(item[1])
    B_N_load_list.append(item[2])
    C_N_load_list.append(item[3])
    ABC_load_list.append(item[4])
    A_B_load_list.append(item[5])
    A_C_load_list.append(item[6])
    B_C_load_list.append(item[7])
    residential.append(item[8])

## For loop to construct the maps from the list of all customers described by our network model to the load objects and phase connection types 
## respectively. Note that only residential customers are included in these maps as charging of commercial/industrial EVs will not be represented
## in this project (assumed that these will charge at larger depots to be built elsewhere).
for idx, PPN in enumerate(customers_PPN_list):
    if residential[idx]:
        app.PrintPlain(idx/len(customers_PPN_list))
        for load in all_LV_loads:
            load_name=load.GetAttribute('loc_name')
            if PPN in load_name:
                if "Bal" in load_name:
                    count=0
                    while count<ABC_load_list[idx]:
                        customers_to_loads_map.append(load)
                        customers_to_phases_map.append("3")
                        count=count+1
                elif "Unbal_AB" in load_name:
                    count=0
                    while count<A_B_load_list[idx]:
                        customers_to_loads_map.append(load)
                        customers_to_phases_map.append("a-b")
                        count=count+1
                elif "Unbal_AC" in load_name:
                    count=0
                    while count<A_C_load_list[idx]:
                        customers_to_loads_map.append(load)
                        customers_to_phases_map.append("a-c")
                        count=count+1
                elif "Unbal_BC" in load_name:
                    count=0
                    while count<B_C_load_list[idx]:
                        customers_to_loads_map.append(load)
                        customers_to_phases_map.append("b-c")
                        count=count+1
                else:
                    count=0
                    while count<A_N_load_list[idx]:
                        customers_to_loads_map.append(load)
                        customers_to_phases_map.append("a")
                        count=count+1
                    count=0
                    while count<B_N_load_list[idx]:
                        customers_to_loads_map.append(load)
                        customers_to_phases_map.append("b")
                        count=count+1
                    count=0
                    while count<C_N_load_list[idx]:
                        customers_to_loads_map.append(load)
                        customers_to_phases_map.append("c")
                        count=count+1

## This while loop reconstructs the array chosen_so_far so that the existing fleet of EVs is accounted for when aiming at the new target total number
## (If no EVs have been encoded into the model yet then chosen_so_far will simply be an array of zeros that is the same length as customers_to_loads_map)
idx=1
map_entries=1
while idx<len(customers_to_loads_map):
    if customers_to_loads_map[idx].GetAttribute('loc_name')!=customers_to_loads_map[idx-1].GetAttribute('loc_name'):
        load=customers_to_loads_map[idx-1]
        load_name=load.GetAttribute('loc_name')
        if "bal_" in load_name:
            local_no_EVs=len(load_name)-load_name.find("bal_")-6
        else:
            local_no_EVs=len(load_name)-load_name.find("al")-2
        for i in range(map_entries):
            chosen_so_far.append(local_no_EVs/map_entries)
        map_entries=1
    else:
        map_entries=map_entries+1
    idx=idx+1
load=customers_to_loads_map[idx-1]
load_name=load.GetAttribute('loc_name')
if "bal_" in load_name:
    local_no_EVs=len(load_name)-load_name.find("bal_")-6
else:
    local_no_EVs=len(load_name)-load_name.find("al")-2
for i in range(map_entries):
    chosen_so_far.append(local_no_EVs/map_entries)

## This while loop encodes the desired number of EVs into the LV loads in the model by appending singular letter characters to the end of the name of
## each LV load to represent EV chargers connected to different phases. The households/customers are selected at random from the customers_to_loads_map.
## The phasing of the EVs are made to correspond to the phase to which the corresponding household/customer is connected; note that many LV loads are 
## representative of multiple households/customers connected to different phases. Note that single phase customers are assumed to use smaller ("size 1")
## 3.68kW EV chargers while three-phase households/customers are assumed to use larger 7.36kW ("size 2") EV chargers.
while sum(chosen_so_far)<(target_residential_avg_EV_count*total_no_households):
    new_EV_created=0
    app.PrintPlain(sum(chosen_so_far)/(target_residential_avg_EV_count*total_no_households))
    x=round(random.uniform(0,len(customers_to_loads_map)-1))
    load=customers_to_loads_map[x]
    load_name=load.GetAttribute('loc_name')
    PPN=load_name[0:load_name.find("_")]        
    if chosen_so_far[x]<upper_EV_limit_per_household:
        if customers_to_phases_map[x]=="a": # Encode a new EV on A phase
            load.SetAttribute('loc_name',load_name+"a")
        elif customers_to_phases_map[x]=="b": # Encode a new EV on b phase
            load.SetAttribute('loc_name',load_name+"b")
        elif customers_to_phases_map[x]=="c": # Encode a new EV on c phase
            load.SetAttribute('loc_name',load_name+"c")
        elif customers_to_phases_map[x]=="3": # Encode a new 3Ph EV
            load.SetAttribute('loc_name',load_name+"3")
        elif customers_to_phases_map[x]=="a-b": # Encode a new a-b EV
            load.SetAttribute('loc_name',load_name+"d")
        elif customers_to_phases_map[x]=="a-c": # Encode a new a-c EV
            load.SetAttribute('loc_name',load_name+"e")
        elif customers_to_phases_map[x]=="b-c": # Encode a new b-c EV
            load.SetAttribute('loc_name',load_name+"f")
    if len(load.GetAttribute('loc_name'))>len(load_name):
        chosen_so_far[x]=chosen_so_far[x]+1

app.PrintPlain(sum(chosen_so_far))            
