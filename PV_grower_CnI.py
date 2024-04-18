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
QDSL_without_export_limits=script.QDSL_without_export_limits

## Construct arrays full of all of a few different types of PowerFactory objects
all_LV_loads = app.GetCalcRelevantObjects('ElmLodlv')
all_QDSLs = app.GetCalcRelevantObjects('ElmQdsl')
all_PVs = app.GetCalcRelevantObjects('ElmPvsys')
all_loads = app.GetCalcRelevantObjects('ElmLod')

## This variables defines the target for the average kVA capacity of PV generation per commercial/industrial customer in the network; note that this should always be set at a value above
## the current average kVA capacity of PV generation per commercial/industrial customer in the model
target_CnI_avg_PV_size = 4.0

total_no_CnI_customers = 785

## Find the current total amount of commercial/industrial PV capacity in the model
Dist_Txs_names_list=['TxSUB_124',	'TxSUB_168',	'TxSUB_3',	'TxSUB_237',	'TxSUB_148',	'TxSUB_58',	'TxSUB_154',	'TxSUB_165',	'TxSUB_31',	'TxSUB_229',	'TxSUB_30',	'TxSUB_240',	'TxSUB_159',	'TxSUB_21',	'TxSUB_234',	'TxSUB_158',	'TxSUB_187',	'TxSUB_26',	'TxSUB_160',	'TxSUB_235',	'TxSUB_208',	'TxSUB_227',	'TxSUB_16',	'TxSUB_182',	'TxSUB_104',	'TxSUB_33',	'TxSUB_105',	'TxSUB_196',	'TxSUB_216',	'TxSUB_218',	'TxSUB_245',	'TxSUB_25',	'TxSUB_179',	'TxSUB_42',	'TxSUB_92',	'TxSUB_123',	'TxSUB_71',	'TxSUB_197',	'TxSUB_41',	'TxSUB_206',	'TxSUB_46',	'TxSUB_39',	'TxSUB_19',	'TxSUB_4',	'TxSUB_191',	'TxSUB_68',	'TxSUB_73',	'TxSUB_204',	'TxSUB_89',	'TxSUB_176',	'TxSUB_5',	'TxSUB_199',	'TxSUB_63',	'TxSUB_137',	'TxSUB_184',	'TxSUB_256',	'TxSUB_67',	'TxSUB_20',	'TxSUB_164',	'TxSUB_163',	'TxSUB_186',	'TxSUB_239',	'TxSUB_180',	'TxSUB_213',	'TxSUB_242',	'TxSUB_189',	'TxSUB_99',	'TxSUB_136',	'TxSUB_95',	'TxSUB_6',	'TxSUB_8',	'TxSUB_220',	'TxSUB_135',	'TxSUB_233',	'TxSUB_205',	'TxSUB_236',	'TxSUB_221',	'TxSUB_243',	'TxSUB_117',	'TxSUB_12',	'TxSUB_60',	'TxSUB_241',	'TxSUB_171',	'TxSUB_207',	'TxSUB_153',	'TxSUB_72',	'TxSUB_174',	'TxSUB_106',	'TxSUB_130',	'TxSUB_177',	'TxSUB_83',	'TxSUB_183',	'TxSUB_32',	'TxSUB_84',	'TxSUB_166',	'TxSUB_173',	'TxSUB_201',	'TxSUB_52',	'TxSUB_157',	'TxSUB_254',	'TxSUB_22',	'TxSUB_226',	'TxSUB_253',	'TxSUB_258',	'TxSUB_47',	'TxSUB_98',	'TxSUB_10',	'TxSUB_252',	'TxSUB_251',	'TxSUB_250',	'TxSUB_45',	'TxSUB_167',	'TxSUB_223',	'TxSUB_140',	'TxSUB_146',	'TxSUB_212',	'TxSUB_34',	'TxSUB_112',	'TxSUB_28',	'TxSUB_101',	'TxSUB_144',	'TxSUB_90',	'TxSUB_14',	'TxSUB_246',	'TxSUB_75',	'TxSUB_214',	'TxSUB_102',	'TxSUB_107',	'TxSUB_161',	'TxSUB_244',	'TxSUB_13',	'TxSUB_211',	'TxSUB_181',	'TxSUB_100',	'TxSUB_7',	'TxSUB_49',	'TxSUB_94',	'TxSUB_169',	'TxSUB_79',	'TxSUB_55',	'TxSUB_27',	'TxSUB_230',	'TxSUB_65',	'TxSUB_215',	'TxSUB_40',	'TxSUB_85',	'TxSUB_121',	'TxSUB_82',	'TxSUB_217',	'TxSUB_134',	'TxSUB_80',	'TxSUB_172',	'TxSUB_2',	'TxSUB_29',	'TxSUB_248',	'TxSUB_97',	'TxSUB_74',	'TxSUB_228',	'TxSUB_88',	'TxSUB_1',	'TxSUB_76',	'TxSUB_38',	'TxSUB_139',	'TxSUB_116',	'TxSUB_56',	'TxSUB_66',	'TxSUB_232',	'TxSUB_87',	'TxSUB_93',	'TxSUB_247',	'TxSUB_249',	'TxSUB_51',	'TxSUB_190',	'TxSUB_231',	'TxSUB_115',	'TxSUB_132',	'TxSUB_122',	'TxSUB_48',	'TxSUB_91',	'TxSUB_209',	'TxSUB_9',	'TxSUB_23',	'TxSUB_222',	'TxSUB_119',	'TxSUB_142',	'TxSUB_178',	'TxSUB_219',	'TxSUB_50',	'TxSUB_11',	'TxSUB_133',	'TxSUB_118',	'TxSUB_162',	'TxSUB_127',	'TxSUB_170',	'TxSUB_96',	'TxSUB_78',	'TxSUB_203',	'TxSUB_103',	'TxSUB_35',	'TxSUB_64',	'TxSUB_194',	'TxSUB_70',	'TxSUB_145',	'TxSUB_128',	'TxSUB_129',	'TxSUB_37',	'TxSUB_113',	'TxSUB_62',	'TxSUB_15',	'TxSUB_151',	'TxSUB_141',	'TxSUB_202',	'TxSUB_143',	'TxSUB_238',	'TxSUB_24',	'TxSUB_225',	'TxSUB_192',	'TxSUB_185',	'TxSUB_200',	'TxSUB_86',	'TxSUB_109',	'TxSUB_120',	'TxSUB_195',	'TxSUB_210',	'TxSUB_17',	'TxSUB_126',	'TxSUB_0',	'TxSUB_44',	'TxSUB_61',	'TxSUB_255',	'TxSUB_81',	'TxSUB_57',	'TxSUB_193',	'TxSUB_188',	'TxSUB_156',	'TxSUB_138',	'TxSUB_54',	'TxSUB_53',	'TxSUB_108',	'TxSUB_111',	'TxSUB_59',	'TxSUB_131',	'TxSUB_175',	'TxSUB_155',	'TxSUB_125',	'TxSUB_224',	'TxSUB_152',	'TxSUB_114',	'TxSUB_43',	'TxSUB_69',	'TxSUB_147',	'TxSUB_110',	'TxSUB_18',	'TxSUB_198',	'TxSUB_260',	'TxSUB_36',	'TxSUB_149',	'TxSUB_77',	'TxSUB_150']
Dis_Txs_residential_list=[0,0,0,	1,	1,	1,	1,	1,	1,	1,	1,	0,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	0,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	0,	1,	0,	0,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	0,	0,	1,	1,	1,	0,	1,	0,	1,	0,	1,	1,	1,	1,	1,	1,	1,	0,	1,	0,	1,	0,	1,	0,	1,	1,	1,	1,	1,	1,	0,	0,	1,	1,	1,	0,	1,	0,	1,	0,	1,	1,	0,	0,	0,	0,	0,	0,	0,	1,	1,	1,	0,	1,	1,	1,	0,	1,	1,	1,	1,	0,	1,	1,	0,	1,	0,	1,	1,	1,	1,	1,	0,	0,	1,	1,	0,	1,	1,	1,	0,	1,	1,	1,	1,	1,	1,	1,	1,	0,	1,	1,	1,	1,	1,	1,	1,	1,	1,	0,	1,	0,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	0,	1,	0,	1,	1,	1,	1,	1,	0,	1,	1,	1,	0,	0,	0,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	0,	1,	1,	1,	1,	0,	1,	1,	1,	1,	1,	0,	1,	1,	1,	0,	1,	1,	1,	0,	1,	1,	1,	1,	1,	1,	1,	0,	1,	1,	0,	1,	1,	1,	1,	1,	1,	1,	1,	1,	1,	0,	1,	1,	1,	0,	0,	1,	1,	1,	1,	0,	1,	1,	1,	0,	1,	0]
total_CnI_PV_capacity = 0
for PV in all_PVs:
    for idx, name in enumerate(Dist_Txs_names_list):
        if not(Dis_Txs_residential_list[idx]):   
            if PV.GetAttribute('bus1').GetAttribute('cterm').GetAttribute('cpSubstat').GetAttribute('loc_name')==name:
                total_CnI_PV_capacity=total_CnI_PV_capacity+PV.GetAttribute('sgn')

app.PrintPlain(total_CnI_PV_capacity)

## Read in some data needed to construct the customers to loads and customers to phases maps
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
added_already=[]
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

## Construct maps from each customer in EQ's records to LV loads in the model and phase connection types respectively. Note that residential loads
## are rejected from these maps
for idx, PPN in enumerate(customers_PPN_list):
    if not(residential[idx]):
        app.PrintPlain(idx/len(customers_PPN_list))
        for load in all_LV_loads:
            load_name=load.GetAttribute('loc_name')
            if load_name.find('Unbal')>0:
                load_PPN=load_name[0:load_name.find('Unbal')]
            else:
                load_PPN=load_name[0:load_name.find('Bal')]                
            if PPN==load_PPN and not Dis_Txs_residential_list[Dist_Txs_names_list.index(load.GetAttribute('bus1').GetAttribute('cterm').GetAttribute('cpSubstat').GetAttribute('loc_name'))]==1:
                if "Unbal" in load_name and "Unbal_" not in load_name:
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
                elif "Bal" in load_name:
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

## Add some additional elements to the customer and phasing maps to account for those commercial/industrial customers that have been represented by the general load components
## (rather than by LV load components)
wb = pd.read_excel('CnI_customers_TxSub_list.xlsx').values
CnI_TxSubs=[]
CnI_TxSub_count=[]
for item in wb:
    CnI_TxSubs.append(item[0])
    CnI_TxSub_count.append(item[1])
for idx, TxSub in enumerate(CnI_TxSubs):
    for load in all_loads:
        if TxSub in load.GetAttribute('bus1').GetAttribute('cterm').GetAttribute('cpSubstat').GetAttribute('loc_name'):
            count=0
            while count<CnI_TxSub_count[idx]:
                customers_to_loads_map.append(load)
                customers_to_phases_map.append("3")
                count=count+1
    
## Add new PV components to the model, or add to the capacity of existing components if they're there, until the total amount of commercial/industrial PV inverter capacity reaches the 
## target that we've established. Note that the new PV capacity is always added randomly to the list of commercial/industrial customers recorded as being connected to the network using 
## a uniform distribution.
while total_CnI_PV_capacity<(target_CnI_avg_PV_size*total_no_CnI_customers):
    all_QDSLs = app.GetCalcRelevantObjects('ElmQdsl')
    all_PVs = app.GetCalcRelevantObjects('ElmPvsys')
    new_PV_created=0
    existing_PV_augmented=0
    app.PrintPlain(total_CnI_PV_capacity/(target_CnI_avg_PV_size*total_no_CnI_customers))
    x=round(random.uniform(0,len(customers_to_loads_map)-1))
    load=customers_to_loads_map[x]
    load_name=load.GetAttribute('loc_name')
    if load_name.find('Unbal')>0:
        PPN=load_name[0:load_name.find('Unbal')]
    else:
        PPN=load_name[0:load_name.find('Bal')]      
    if customers_to_phases_map[x]=="a": # Make a new PV on A phase
        if load.GetAttribute('slinir')>0:
            for PV in all_PVs:
                if PV.GetAttribute('loc_name')==PPN+"_pv_A" or PV.GetAttribute('loc_name')==PPN+"_future_pv_A":
                    existing_capacity=PV.GetAttribute('sgn')
                    PV.SetAttribute('sgn',8+existing_capacity)
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8
                    new_PV_name=PV.GetAttribute('loc_name')
                    new_PV=PV            
                    existing_PV_augmented=1
            if existing_PV_augmented==0:
                cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                new_PV_name=PPN+'_future_pv_A'
                new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                new_PV_name=new_PV.GetAttribute('loc_name')
                new_PV.SetAttribute('phtech',3)
                new_PV.SetAttribute('cosn',1)
                new_PV.SetAttribute('bus1',cubic)
                new_PV.SetAttribute('sgn',8)
                new_PV.GetAttribute('bus1').SetAttribute('it2p1',0)
                new_PV_created=1
                total_CnI_PV_capacity=total_CnI_PV_capacity+8
    elif customers_to_phases_map[x]=="b": # Make a new PV on b phase
        if load.GetAttribute('slinis')>0:
            for PV in all_PVs:
                if PV.GetAttribute('loc_name')==PPN+"_pv_B" or PV.GetAttribute('loc_name')==PPN+"_future_pv_B":
                    existing_capacity=PV.GetAttribute('sgn')
                    PV.SetAttribute('sgn',8+existing_capacity)
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8   
                    new_PV_name=PV.GetAttribute('loc_name')     
                    new_PV=PV            
                    existing_PV_augmented=1
            if existing_PV_augmented==0:
                cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                new_PV_name=PPN+'_future_pv_B'
                new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                new_PV_name=new_PV.GetAttribute('loc_name')
                new_PV.SetAttribute('phtech',3)
                new_PV.SetAttribute('cosn',1)
                new_PV.SetAttribute('bus1',cubic)
                new_PV.SetAttribute('sgn',8)
                new_PV.GetAttribute('bus1').SetAttribute('it2p1',1)
                new_PV_created=1
                total_CnI_PV_capacity=total_CnI_PV_capacity+8
    elif customers_to_phases_map[x]=="c": # Make a new PV on c phase
        if load.GetAttribute('slinit')>0:
            for PV in all_PVs:
                if PV.GetAttribute('loc_name')==PPN+"_pv_C" or PV.GetAttribute('loc_name')==PPN+"_future_pv_C":
                    existing_capacity=PV.GetAttribute('sgn')
                    PV.SetAttribute('sgn',8+existing_capacity)
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8    
                    new_PV_name=PV.GetAttribute('loc_name')               
                    new_PV=PV
                    existing_PV_augmented=1
            if existing_PV_augmented==0:
                cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                new_PV_name=PPN+'_future_pv_C'
                new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                new_PV_name=new_PV.GetAttribute('loc_name')
                new_PV.SetAttribute('phtech',3)
                new_PV.SetAttribute('cosn',1)
                new_PV.SetAttribute('bus1',cubic)
                new_PV.SetAttribute('sgn',8)
                new_PV.GetAttribute('bus1').SetAttribute('it2p1',2)
                new_PV_created=1
                total_CnI_PV_capacity=total_CnI_PV_capacity+8
    elif customers_to_phases_map[x]=="3": # Make a new 3Ph PV            
        if load.GetAttribute('slini')>0:
            for PV in all_PVs:
                if PV.GetAttribute('loc_name')==PPN+"_pv_3Ph" or PV.GetAttribute('loc_name')==PPN+"_future_pv_3Ph":
                    existing_capacity=PV.GetAttribute('sgn')
                    PV.SetAttribute('sgn',8+existing_capacity)
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8
                    new_PV_name=PV.GetAttribute('loc_name')
                    new_PV=PV
                    existing_PV_augmented=1
            if existing_PV_augmented==0:
                cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                new_PV_name=PPN+'_future_pv_3Ph'
                new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                new_PV_name=new_PV.GetAttribute('loc_name')
                new_PV.SetAttribute('phtech',0)
                new_PV.SetAttribute('bus1',cubic)
                new_PV.SetAttribute('sgn',8)
                new_PV_created=1
                total_CnI_PV_capacity=total_CnI_PV_capacity+8
    elif customers_to_phases_map[x]=="a-b": # Make a new A or B phase PV            
        y=random.uniform(0,1)
        if y<0.5:
            if load.GetAttribute('slinir')>0:
                for PV in all_PVs:
                    if PV.GetAttribute('loc_name')==PPN+"_pv_A" or PV.GetAttribute('loc_name')==PPN+"_future_pv_A":
                        existing_capacity=PV.GetAttribute('sgn')
                        PV.SetAttribute('sgn',8+existing_capacity)
                        total_CnI_PV_capacity=total_CnI_PV_capacity+8
                        new_PV_name=PV.GetAttribute('loc_name')
                        new_PV=PV            
                        existing_PV_augmented=1
                if existing_PV_augmented==0:
                    cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                    new_PV_name=PPN+'_future_pv_A'
                    new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                    new_PV_name=new_PV.GetAttribute('loc_name')
                    new_PV.SetAttribute('phtech',3)
                    new_PV.SetAttribute('cosn',1)
                    new_PV.SetAttribute('bus1',cubic)
                    new_PV.SetAttribute('sgn',8)
                    new_PV.GetAttribute('bus1').SetAttribute('it2p1',0)
                    new_PV_created=1
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8
        else:
            if load.GetAttribute('slinir')>0:
                for PV in all_PVs:
                    if PV.GetAttribute('loc_name')==PPN+"_pv_B" or PV.GetAttribute('loc_name')==PPN+"_future_pv_B":
                        existing_capacity=PV.GetAttribute('sgn')
                        PV.SetAttribute('sgn',8+existing_capacity)
                        total_CnI_PV_capacity=total_CnI_PV_capacity+8
                        new_PV_name=PV.GetAttribute('loc_name')
                        new_PV=PV            
                        existing_PV_augmented=1
                if existing_PV_augmented==0:
                    cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                    new_PV_name=PPN+'_future_pv_B'
                    new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                    new_PV_name=new_PV.GetAttribute('loc_name')
                    new_PV.SetAttribute('phtech',3)
                    new_PV.SetAttribute('cosn',1)
                    new_PV.SetAttribute('bus1',cubic)
                    new_PV.SetAttribute('sgn',8)
                    new_PV.GetAttribute('bus1').SetAttribute('it2p1',1)
                    new_PV_created=1
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8
    elif customers_to_phases_map[x]=="a-c": # Make a new A or C phase PV            
        y=random.uniform(0,1)
        if y<0.5:
            if load.GetAttribute('slinir')>0:
                for PV in all_PVs:
                    if PV.GetAttribute('loc_name')==PPN+"_pv_A" or PV.GetAttribute('loc_name')==PPN+"_future_pv_A":
                        existing_capacity=PV.GetAttribute('sgn')
                        PV.SetAttribute('sgn',8+existing_capacity)
                        total_CnI_PV_capacity=total_CnI_PV_capacity+8
                        new_PV_name=PV.GetAttribute('loc_name')
                        new_PV=PV            
                        existing_PV_augmented=1
                if existing_PV_augmented==0:
                    cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                    new_PV_name=PPN+'_future_pv_A'
                    new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                    new_PV_name=new_PV.GetAttribute('loc_name')
                    new_PV.SetAttribute('phtech',3)
                    new_PV.SetAttribute('cosn',1)
                    new_PV.SetAttribute('bus1',cubic)
                    new_PV.SetAttribute('sgn',8)
                    new_PV.GetAttribute('bus1').SetAttribute('it2p1',0)
                    new_PV_created=1
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8
        else:
            if load.GetAttribute('slinir')>0:
                for PV in all_PVs:
                    if PV.GetAttribute('loc_name')==PPN+"_pv_C" or PV.GetAttribute('loc_name')==PPN+"_future_pv_C":
                        existing_capacity=PV.GetAttribute('sgn')
                        PV.SetAttribute('sgn',8+existing_capacity)
                        total_CnI_PV_capacity=total_CnI_PV_capacity+8
                        new_PV_name=PV.GetAttribute('loc_name')
                        new_PV=PV            
                        existing_PV_augmented=1
                if existing_PV_augmented==0:
                    cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                    new_PV_name=PPN+'_future_pv_C'
                    new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                    new_PV_name=new_PV.GetAttribute('loc_name')
                    new_PV.SetAttribute('phtech',3)
                    new_PV.SetAttribute('cosn',1)
                    new_PV.SetAttribute('bus1',cubic)
                    new_PV.SetAttribute('sgn',8)
                    new_PV.GetAttribute('bus1').SetAttribute('it2p1',2)
                    new_PV_created=1
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8
    elif customers_to_phases_map[x]=="b-c": # Make a new B or C phase PV            
        y=random.uniform(0,1)
        if y<0.5:
            if load.GetAttribute('slinir')>0:
                for PV in all_PVs:
                    if PV.GetAttribute('loc_name')==PPN+"_pv_B" or PV.GetAttribute('loc_name')==PPN+"_future_pv_B":
                        existing_capacity=PV.GetAttribute('sgn')
                        PV.SetAttribute('sgn',8+existing_capacity)
                        total_CnI_PV_capacity=total_CnI_PV_capacity+8
                        new_PV_name=PV.GetAttribute('loc_name')
                        new_PV=PV            
                        existing_PV_augmented=1
                if existing_PV_augmented==0:
                    cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                    new_PV_name=PPN+'_future_pv_B'
                    new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                    new_PV_name=new_PV.GetAttribute('loc_name')
                    new_PV.SetAttribute('phtech',3)
                    new_PV.SetAttribute('cosn',1)
                    new_PV.SetAttribute('bus1',cubic)
                    new_PV.SetAttribute('sgn',8)
                    new_PV.GetAttribute('bus1').SetAttribute('it2p1',1)
                    new_PV_created=1
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8
        else:
            if load.GetAttribute('slinir')>0:
                for PV in all_PVs:
                    if PV.GetAttribute('loc_name')==PPN+"_pv_C" or PV.GetAttribute('loc_name')==PPN+"_future_pv_C":
                        existing_capacity=PV.GetAttribute('sgn')
                        PV.SetAttribute('sgn',8+existing_capacity)
                        total_CnI_PV_capacity=total_CnI_PV_capacity+8
                        new_PV_name=PV.GetAttribute('loc_name')
                        new_PV=PV            
                        existing_PV_augmented=1
                if existing_PV_augmented==0:
                    cubic=load.GetAttribute('bus1').GetAttribute('cterm').CreateObject('StaCubic')
                    new_PV_name=PPN+'_future_pv_C'
                    new_PV=Grid.CreateObject('ElmPvsys',new_PV_name)            
                    new_PV_name=new_PV.GetAttribute('loc_name')
                    new_PV.SetAttribute('phtech',3)
                    new_PV.SetAttribute('cosn',1)
                    new_PV.SetAttribute('bus1',cubic)
                    new_PV.SetAttribute('sgn',8)
                    new_PV.GetAttribute('bus1').SetAttribute('it2p1',2)
                    new_PV_created=1
                    total_CnI_PV_capacity=total_CnI_PV_capacity+8
    if new_PV_created:
        new_QDSL=Grid.CreateObject('ElmQdsl','QDSL_'+new_PV_name)
        new_QDSL.typ_id=QDSL_without_export_limits
        new_QDSL.initVals=[1000,2,1,1]
        new_QDSL.objects=[new_PV]
    else:
        QDSL_existing=0
        for QDSL in all_QDSLs:
            if "QDSL_"+new_PV_name==QDSL.GetAttribute('loc_name'):
                QDSL.typ_id=QDSL_without_export_limits
                QDSL.initVals=[1000,2,1,1]
                QDSL_existing=1
        if QDSL_existing==0:
            new_QDSL=Grid.CreateObject('ElmQdsl','QDSL_'+new_PV_name)
            new_QDSL.typ_id=QDSL_without_export_limits
            new_QDSL.initVals=[1000,2,1,1]
            new_QDSL.objects=[new_PV]

