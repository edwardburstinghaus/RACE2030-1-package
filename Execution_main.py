import powerfactory as pf
import importlib
import exec_utils as execUtils
importlib.reload(execUtils)
import pandas as pd
import numpy as np
import datetime
from openpyxl import Workbook
import random
import os
import datetime
import winsound
winsound.Beep(440,500)
app = pf.GetApplication()
app.ClearOutputWindow()
app.EchoOn()
script = app.GetCurrentScript()
baselined_model_PVs = script.baselined_model_PVs
grid = script.Grid
baselined_model_PVs = execUtils.extractObjectsFromSet(baselined_model_PVs)

## Construct arrays full of all of a few different types of PowerFactory objects
all_loads = app.GetCalcRelevantObjects('ElmLod')
all_LV_loads = app.GetCalcRelevantObjects('ElmLodlv')
all_buses = app.GetCalcRelevantObjects('ElmTerm')
all_lines = app.GetCalcRelevantObjects('ElmLne')
all_QDSLs = app.GetCalcRelevantObjects('ElmQdsl')
all_transformers = app.GetCalcRelevantObjects('ElmTr2')
all_PVs = app.GetCalcRelevantObjects('ElmPvsys')       
all_coups = app.GetCalcRelevantObjects('ElmCoup')
all_subs = app.GetCalcRelevantObjects('ElmTrfstat')
time1 = datetime.datetime.now()

## Variables to configure what Season you want to simulate, whether you want PV on and whether you want QDSLs on
Season = "Winter"
SolarOn = 1
QDSLsOn = 1
EVsOn = 0

## Read in the peak PV utilisation factors for PV generators that were in the model at the baselining stage
## for which we have generation curves that we've derived from transformer monitor data
cwd = os.getcwd()
wb = pd.read_excel(cwd+"\\Peak pgini sheets\\"+Season+'_peak_pgini_values.xlsx').values
baselined_year_peak_solar_outputs=[] 
baselined_year_solar_capacities=[]
peak_solar_outputs_names_array=[]
for item in wb:
    baselined_year_peak_solar_outputs.append(item[0])
    baselined_year_solar_capacities.append(item[1])
    peak_solar_outputs_names_array.append(item[2])

if EVsOn:
    ## Construct the maps from notional EVs to LV loads
    EVs_to_loads_map1, EVs_to_phases_map1, EVs_to_loads_map2, total_size1_charger_EVs, total_size2_charger_EVs = execUtils.construct_EV_maps(all_LV_loads,app)
    if total_size1_charger_EVs+total_size2_charger_EVs==0:
        EVsOn=0

## Configure a series of arrays that facilitate the EV charging functionality
Electric_Nation_Weekday_Level1_profile = [0.403614,	0.26506,	0.150602,	0.096386,	0.066265,	0.036145,	0.024096,	0,	0,	0,	0,	0,	0,	0,	0,	0.03012,	0.048193,	0.078313,	0.138554,	0.156627,	0.162651,	0.180723,	0.168675,	0.174699,	0.198795,	0.259036,	0.289157,	0.307229,	0.337349,	0.325301,	0.343373,	0.421687,	0.493976,	0.560241,	0.638554,	0.843373,	0.957831,	1.024096,	1.090361,	1.090361,	1.072289,	1.054217,	1.036145,	1,	0.927711,	0.831325,	0.692771,	0.53012]
Electric_Nation_Weekday_Level2_profile = [1.2411,	1.038575,	0.8511,	0.64895,	0.4255,	0.31025,	0.2411,	0.181828571,	0.141172727,	0.1064,	0,	0.033471429,	0.022119403,	0.057044776,	0.112923077,	0.197416667,	0.297916667,	0.394102857,	0.5106,	0.6099125,	0.620042857,	0.62552,	0.622809091,	0.666233333,	0.6879,	0.71940625,	0.73538125,	0.785714286,	0.844,	0.8865,	0.870161538,	0.964035897,	1.025058824,	1.095324,	1.224290476,	1.487719048,	1.727617647,	1.902808824,	2.078,	2.155510256,	2.1915,	2.1915,	2.164114286,	2.2128,	2.049653333,	1.8936,	1.7518,	1.535917391]
Electric_Nation_Weekday_Level1_profile = np.array(Electric_Nation_Weekday_Level1_profile)
Electric_Nation_Weekday_Level2_profile = np.array(Electric_Nation_Weekday_Level2_profile)
EV_curve1 =Electric_Nation_Weekday_Level1_profile/3.68
EV_curve2 =Electric_Nation_Weekday_Level2_profile/7.36
EVs_charging_list1=[]
EVs_charging_list2=[]
current_EV_charging_count1=0
current_EV_charging_count2=0

## Code to access the scaling factors which realise the baselined LV load sizes
if SolarOn:
    dataFilename="Load_curve_"+Season+" - baselined for high solar.xlsx"
else:
    dataFilename="Load_curve_"+Season+" - baselined for low solar.xlsx"
dataFilepath = cwd + "\\Scaling sheets"
dataFilepath = execUtils.ConcatFilePathFileName(dataFilepath,dataFilename)
solarProfile = execUtils.loadSolarProfile(dgApp=app, excelPath=dataFilepath)
todProfile = execUtils.loadTodProfile(dgApp=app, excelPath=dataFilepath)
subLoadProfile = execUtils.loadSubLoadProfile(dgApp=app, excelPath=dataFilepath)
matchingSubstations, err = execUtils.checkDataFile(dgApp=app, excelPath=dataFilepath)
subLoadProfile = execUtils.loadSubLoadProfile(dgApp=app, excelPath=dataFilepath)
grid_source = app.GetCalcRelevantObjects('ElmXnet')[0]

## Obtain the load flow object
ldf = app.GetFromStudyCase("ComLdf")

## Configure PV components
for idx, PV in enumerate(all_PVs):
    PV.SetAttribute('scale0',1.0)
    if SolarOn:
        PV.SetAttribute('outserv',0)
    else:
        PV.SetAttribute('outserv',1)
        PV.SetAttribute('pgini',0.0)

# Configure QDSL models
for QDSL in all_QDSLs:
    if QDSLsOn:
        QDSL.SetAttribute('outserv',0)
    else:
        QDSL.SetAttribute('outserv',1)
        
## List of LV buses to have STATCOMs placed at them
# bus_list=['B1156','B4424']

## Implement STATCOMs
# all_STATCOMs=execUtils.realise_STATCOMs(all_buses,bus_list,grid,ldf)

# Set up some arrays
bus_levels, bus_TODs, bus_names, bus_A_phase_voltages, bus_B_phase_voltages, bus_C_phase_voltages, bus_N_voltages, bus_A_phase_voltage_angles, bus_B_phase_voltage_angles, bus_C_phase_voltage_angles, bus_N_voltage_angles, bus_pos_seq_voltage_magnitudes, bus_neg_seq_voltage_magnitudes, bus_avg_voltage_magnitudes, substation_names, transformer_names, transformer_avg_ln_voltage_mags, transformer_avg_lg_voltage_mags, transformer_pos_seq_voltage_mags, transformer_active_power_flows, transformer_reactive_power_flows, transformer_TODs, transformer_loadings, line_names, line_levels, line_TODs, line_loadings, line_substations, line_types = execUtils.prepareArrays()

## Code for a full-day run
# for todIdx, tod in enumerate(todProfile):
#     app.PrintPlain(tod)

# Code for isolating one time of day (noon) without having to remove the indent for the for loop the facilitates iteration over more than one time point in the day in general
shortlist=[todProfile[24]]
for todIdx, tod in enumerate(shortlist):
    todIdx=24

    ldf.SetAttribute('nsteps',1)

    # Set the loads in the substation
    total_size1_EVs=0
    total_size2_EVs=0
    total_size1_EVsOn=0
    total_size2_EVsOn=0
    for subIdx, sub in enumerate(matchingSubstations):
        subLoadScaler = subLoadProfile[subIdx][todIdx]
        if subLoadScaler>0:
            subLoadScaler=subLoadScaler
        else:
            subLoadScaler=subLoadScaler


        # Set the load scaling to the value
        objs = sub.GetContents('*.ElmLod')
        for obj in objs:
            obj.SetAttribute('scale0', subLoadScaler)

        # Set the lv loads created in the specific grids
        objs = sub.GetContents('*.ElmLodlv')
        for obj in objs:
            ## Rescale (i.e. alter the scale0 parameter) the portion of each of the LV loads in the model that is not attributable to EVs as per
            ## the original baselined load curves for each season and set of solar irradiance conditions, in such a way that the portion of each
            ## LV load that represents power flowing to an EV or EVs is kept constant. Note that if set_EVs_charging() is commented out than the 
            ## latter portion will be zero in general and this function will simply rescale the loads as per the curves in the baselining sheets.  
            total_size1_EVs, total_size2_EVs, total_size1_EVsOn, total_size2_EVsOn = execUtils.rescale_non_EV_LV_load_portion(obj,subLoadScaler,total_size1_EVs, total_size2_EVs, total_size1_EVsOn, total_size2_EVsOn,app)

    if EVsOn:
        ## Activation of EV charging
        EVs_charging_list1, EVs_charging_list2, current_EV_charging_count1, current_EV_charging_count2  = execUtils.set_EVs_charging(EVs_charging_list1, EVs_charging_list2, EVs_to_loads_map1, EVs_to_phases_map1, EVs_to_loads_map2, total_size1_charger_EVs, total_size2_charger_EVs, EV_curve1, EV_curve2, todIdx, current_EV_charging_count1, current_EV_charging_count2, app)

    ## Initialise pgini (initial active power setpoint) values for each of the PV components in the model; note that the QDSLs may alter the 
    ## effective active power that is generated by some of the PV components (i.e. as a result of curtailment)
    if SolarOn:
        for idx, PV in enumerate(baselined_model_PVs):
            capacity = PV.GetAttribute('sgn')
            if PV.GetAttribute('loc_name')!=peak_solar_outputs_names_array[idx]:
                app.PrintPlain("ERROR: Order of baselined_model_PVs set in model doesn't match pgini spreadsheet")
                break
            PV.SetAttribute('pgini',capacity*baselined_year_peak_solar_outputs[idx]/baselined_year_solar_capacities[idx]*solarProfile[todIdx])
        for PV in all_PVs:
            capacity = PV.GetAttribute('sgn')
            if "future" in PV.GetAttribute('loc_name'):
                if Season=="Summer" or Season=="Spring":
                    PV.SetAttribute('pgini',capacity*1.0*solarProfile[todIdx]/max(solarProfile))
                elif Season=="Winter":
                    PV.SetAttribute('pgini',capacity*0.8289304516*solarProfile[todIdx]/max(solarProfile))
                elif Season=="Autumn":
                    PV.SetAttribute('pgini',capacity*0.9762955505*solarProfile[todIdx]/max(solarProfile))


    ## This line takes a guess at what the outcome of the zone substation LDC algorithm will be before solving the load flow for the network and checking the network total
    ## active power flow; use only when simulating limited time-periods in the day for which this prediction is straightforward e.g. for simulation at noon under high solar
    ## conditions, 0.9875pu is a reliable prediction for the zone substation effective MV votlage source value. Comment this line out for faster performance
    ## When running a simulation across a whole day
    grid_source.SetAttribute('usetp',0.9875)
    
    err = ldf.Execute()       
    
    ## Check if the load flow needs to be repeated due to operation of the zone substation LDC algorithms 
    grid_source_P_flow=grid_source.GetAttribute('m:Psum:bus1')
    grid_source_voltage_setpoint=round(grid_source.GetAttribute('usetp'),4)
    rerun=0
    # if grid_source_P_flow>30643.2:
    #     if grid_source_voltage_setpoint!=1.0125:
    #         rerun=1
    #         grid_source.SetAttribute('usetp',1.0125)
    # elif grid_source_P_flow<14968.2:
    #     if grid_source_voltage_setpoint!=0.9875:
    #         rerun=1
    #         grid_source.SetAttribute('usetp',0.9875)
    # else:
    #     if grid_source_voltage_setpoint!=1.0:
    #         rerun=1
    #         grid_source.SetAttribute('usetp',1.0)
    # if rerun:
    #     app.PrintPlain("Rerunning...")
    #     err = ldf.Execute()                 

    ## Print useful information about the active and reactive power flows in the current simulation
    grid_source_P_flow=grid_source.GetAttribute('m:Psum:bus1')
    grid_source_voltage_setpoint=round(grid_source.GetAttribute('usetp'),4)
    app.PrintPlain("Zone sub total kW flow is: "+str(grid_source_P_flow))
    grid_source_Q_flow=grid_source.GetAttribute('m:Qsum:bus1') 
    app.PrintPlain("Zone sub total kVAr flow is: "+str(grid_source_Q_flow))
    current_total_load=0
    for load in all_LV_loads:
        current_total_load=current_total_load+load.GetAttribute('slini')*load.GetAttribute('scale0')
    # for load in all_loads:
    #     current_total_load=current_total_load+load.GetAttribute('slini')*load.GetAttribute('scale0')
    app.PrintPlain("Total kVA of all loads is: "+str(current_total_load))
    total_PV_generation=0
    for PV in all_PVs:
        if PV.GetAttribute('bus1').GetAttribute('cterm').GetAttribute('uknom')<10:
            total_PV_generation=total_PV_generation+PV.GetAttribute('s:pset')*1000
    app.PrintPlain('Total kW being generated by all PVs is: '+str(total_PV_generation))

    ## Save a spreadsheet on what the curtailment distribution was and how accurate it was
    if QDSLsOn:
        execUtils.save_curtailment_results(app, all_QDSLs, str(todIdx))
    
    ### Code that facilitates the storing of output data in various Excel sheets
    # bus_levels, bus_names, bus_TODs, bus_A_phase_voltages, bus_B_phase_voltages, bus_C_phase_voltages, bus_N_voltages, bus_A_phase_voltage_angles, bus_B_phase_voltage_angles, bus_C_phase_voltage_angles, bus_N_voltage_angles, bus_pos_seq_voltage_magnitudes, bus_neg_seq_voltage_magnitudes, bus_avg_voltage_magnitudes, substation_names = execUtils.pack_bus_results(QDSLsOn, SolarOn, Season, app, todIdx, bus_levels, bus_names, bus_TODs, bus_A_phase_voltages, bus_B_phase_voltages, bus_C_phase_voltages, bus_N_voltages, bus_A_phase_voltage_angles, bus_B_phase_voltage_angles, bus_C_phase_voltage_angles, bus_N_voltage_angles, bus_pos_seq_voltage_magnitudes, bus_neg_seq_voltage_magnitudes, bus_avg_voltage_magnitudes, all_buses, tod, substation_names)
    # transformer_names, transformer_TODs, transformer_avg_ln_voltage_mags, transformer_active_power_flows, transformer_reactive_power_flows, transformer_loadings = execUtils.pack_transformer_results(transformer_names, transformer_TODs, transformer_avg_ln_voltage_mags, transformer_active_power_flows, transformer_reactive_power_flows, transformer_loadings, all_transformers, tod)            
    # line_names, line_TODs, line_levels, line_loadings, line_substations, line_types = execUtils.pack_line_results(line_names, line_TODs, line_levels, line_loadings, all_lines, tod, line_substations, line_types)    
# execUtils.save_bus_results(todIdx, bus_levels, bus_TODs, QDSLsOn, SolarOn, bus_names, bus_A_phase_voltages, bus_B_phase_voltages, bus_C_phase_voltages, bus_N_voltages, bus_A_phase_voltage_angles, bus_B_phase_voltage_angles, bus_C_phase_voltage_angles, bus_N_voltage_angles, bus_pos_seq_voltage_magnitudes, bus_neg_seq_voltage_magnitudes, bus_avg_voltage_magnitudes, Season, app, substation_names)
# execUtils.save_transformer_results(transformer_names, app, transformer_TODs, QDSLsOn, SolarOn, transformer_avg_ln_voltage_mags, transformer_active_power_flows, transformer_reactive_power_flows, transformer_loadings, Season)
# execUtils.save_line_results(app, line_TODs, QDSLsOn, SolarOn, line_levels, line_names, line_loadings, Season, line_substations, line_types)

if EVsOn:
    ## De-active all EVs in the model; note that if this function does not run corretly at the end of a simulation when EVs were on, the sizes of the LV loads will be configured 
    ## incorrectly; to correct them either run the "LV_Loads_Reset" script with the variable "Season" set to "Winter" or re-load an older saved version of the model
    execUtils.deactivate_all_EVs(EVs_charging_list1, EVs_charging_list2, EVs_to_loads_map1, EVs_to_phases_map1, EVs_to_loads_map2, app)
    EVs_charging_list1=[]
    EVs_charging_list2=[]
    current_EV_charging_count1=0
    current_EV_charging_count2=0
    no_EVs_charging=0
    for load in all_LV_loads:
        load_name=load.GetAttribute('loc_name')
        if "bal_" in load_name:
            EV_substring=load_name[load_name.find("bal_")+6:len(load_name)]
        else:
            EV_substring=load_name[load_name.find("al")+2:len(load_name)]
        for i in EV_substring:
            if i.isupper():
                no_EVs_charging=no_EVs_charging+1

## Delete all STATCOM components that were made for this simulation
# for STATCOM in all_STATCOMs:
#     STATCOM.Delete()

time2 = datetime.datetime.now()
app.PrintPlain("Duration of simulations was "+str(time2-time1))
winsound.Beep(440,500)
app.EchoOn()