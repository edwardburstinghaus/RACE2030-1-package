import powerfactory as pf
import pandas as pd

app = pf.GetApplication()
app.ClearOutputWindow()
script = app.GetCurrentScript()
grid = script.Grid
LV_95_ABC = script.LV_95_ABC

LV_OH_spans_scheduled_for_upgrade=0
LV_buried_cable_runs_scheduled_for_upgrade=0
MV_OH_line_upgrades_idx=0
MV_UG_line_upgrades_idx=0
LV_OH_line_upgrades_idx=0
LV_UG_line_upgrades_idx=0
Tx_upgrades_idx=0
MOON_length=0
PLUTO_length=0
LV_Double_95_ABC_length=0
LV_300_Al_3p5C_length=0
LV_240_Al_3C_NS_length=0
LV_240_Al_4C_XLPEPVC_length=0
LV_25_Cu_4C_PLY_HDPE_PVC_length=0
LV_25_Cu_4C_XLPE_length=0
LV_70_Cu_3p5C_length=0
LV_70_Cu_4C_HDPE_length=0
LV_150_Cu_3C_NS_length=0
LV_120_Cu_4C_XLPEPVC_length=0
LV_240_Cu_4C_XLPEPVC_length=0

# pass_idx controls which pass at the upgrades for a given model we're applying (i.e. multiple passes are necessary to prevent all thermal overloads)
pass_idx=1

#### LV_types: all the conductor and tower types for LV lines. This set is pre-defined within the model.
#### MV_types: all the conductor and line types for MV lines
LV_types = []
MV_types = []

all_cable_types = app.GetCalcRelevantObjects('TypCabsys')
all_tower_types = app.GetCalcRelevantObjects('TypTow')
for item in all_cable_types:
    LV_types.append(item)
for item in all_tower_types:
    LV_types.append(item)

all_conductor_types = app.GetCalcRelevantObjects('TypCon')
all_line_types = app.GetCalcRelevantObjects('TypLne')
for item in all_conductor_types:
    if item.GetAttribute('fold_id').GetAttribute('loc_name') == "Less Used OH Types" or item.GetAttribute('fold_id').GetAttribute('loc_name') == "Common OH Types":
        MV_types.append(item)
for item in all_line_types:
    if item.GetAttribute('fold_id').GetAttribute('loc_name') == "SEQ Standard Types":
        MV_types.append(item)

#### all_lines: all the lines within the model
all_lines = app.GetCalcRelevantObjects('ElmLne')

#### LV_feeder_set_all: the LV feeders of which all lines need to be updated
#### LV_feeder_set_OHL: the LV feeders of which only over head lines need to be updated
#### LV_feeder_set_cable: the LV feeders of which only underground cables need to be updated
#### LV_Line_string_set: all the LV lines need to be upgrade
#### MV_Line_string_set: all the MV lines need to be upgrade
LV_feeder_set_all = ['TxSUB_0','TxSUB_1']
LV_feeder_set_cable = []
LV_feeder_set_OHL=[]
LV_line_string_set=[]
MV_line_string_set = []
transformers_to_be_upgraded=[]

LV_line_set=[]
MV_line_set=[]
for idx, line in enumerate(all_lines):
    if line.GetAttribute('loc_name') in MV_line_string_set:
        MV_line_set.append(line)
    elif line.GetAttribute('loc_name') in LV_line_string_set:
        LV_line_set.append(line)

for item in all_lines:
    try:
        line_name = item.GetAttribute('c_ptow').GetAttribute('loc_name')
        if item.GetAttribute('bus1').GetAttribute('cterm').GetAttribute('cpSubstat').GetAttribute('loc_name') in LV_feeder_set_all:
            LV_line_set.append(item)
            if "OHL" in line_name:
                LV_OH_spans_scheduled_for_upgrade=LV_OH_spans_scheduled_for_upgrade+1
            else:
                LV_buried_cable_runs_scheduled_for_upgrade=LV_buried_cable_runs_scheduled_for_upgrade+1
        elif item.GetAttribute('bus1').GetAttribute('cterm').GetAttribute('cpSubstat').GetAttribute('loc_name') in LV_feeder_set_OHL:
            if "OHL" in line_name:
                LV_line_set.append(item)
                LV_OH_spans_scheduled_for_upgrade=LV_OH_spans_scheduled_for_upgrade+1
            else:
                continue
        elif item.GetAttribute('bus1').GetAttribute('cterm').GetAttribute('cpSubstat').GetAttribute('loc_name') in LV_feeder_set_cable:
            if "cable" in line_name:
                LV_line_set.append(item)
                LV_buried_cable_runs_scheduled_for_upgrade=LV_buried_cable_runs_scheduled_for_upgrade+1
            else:
                continue
    except:
        continue

#### Upgrade the LV lines to the next level
for Line_item in LV_line_set:
    line_name = Line_item.GetAttribute('c_ptow').GetAttribute('loc_name')

    #### Upgrade overhead lines
    if "OHL" in line_name:
        line_type_name = Line_item.GetAttribute('c_ptow').GetAttribute('pGeo:0').GetAttribute('loc_name')
        if line_type_name == "LV - 7/14 Cu" or line_type_name == "LV - MARS" or line_type_name == "LV - MINK" or line_type_name == "LV - 7/.104" or line_type_name == "LV - BANANA" or line_type_name == "LV - 19/.083" or line_type_name == "LV - LIBRA":
            new_type = "LV - MOON"
            MOON_length=MOON_length+Line_item.GetAttribute('dline')/1000
        elif line_type_name == "LV - MOON":
            new_type = "LV - PLUTO"
            PLUTO_length=PLUTO_length+Line_item.GetAttribute('dline')/1000
        elif line_type_name == "LV - 95 ABC":
            existing_coupling_obj=Line_item.GetAttribute('c_ptow')
            line1=existing_coupling_obj.GetAttribute('plines:0')            
            line1_name=line1.GetAttribute('loc_name')
            name=Line_item.GetAttribute('loc_name')
            new_type = "LV - Double 95 ABC"
            if line1_name[len(line1_name)-2:len(line1_name)]!='_u' and line1_name[len(line1_name)-3:len(line1_name)-1]!='2_':
                line2=existing_coupling_obj.GetAttribute('plines:1')            
                line3=existing_coupling_obj.GetAttribute('plines:2')            
                line4=existing_coupling_obj.GetAttribute('plines:3')
                line2_name=line2.GetAttribute('loc_name')
                line3_name=line3.GetAttribute('loc_name')
                line4_name=line4.GetAttribute('loc_name')
                line1.SetAttribute('loc_name',line1_name+"_u")
                line2.SetAttribute('loc_name',line2_name+"_u")
                line3.SetAttribute('loc_name',line3_name+"_u")
                line4.SetAttribute('loc_name',line4_name+"_u")
                line5=grid.CreateObject('ElmLne',name+'2_A')
                line6=grid.CreateObject('ElmLne',name+'2_B')
                line7=grid.CreateObject('ElmLne',name+'2_C')
                line8=grid.CreateObject('ElmLne',name+'2_N')
                length=Line_item.GetAttribute('dline')
                line5.SetAttribute('dline',length)
                line6.SetAttribute('dline',length)
                line7.SetAttribute('dline',length)
                line8.SetAttribute('dline',length)
                bus1=Line_item.GetAttribute('bus1').GetAttribute('cterm')
                cubic_A=bus1.CreateObject('StaCubic')
                line5.SetAttribute('bus1',cubic_A)
                cubic_B=bus1.CreateObject('StaCubic')
                line6.SetAttribute('bus1',cubic_B)
                cubic_C=bus1.CreateObject('StaCubic')
                line7.SetAttribute('bus1',cubic_C)
                cubic_N=bus1.CreateObject('StaCubic')
                line8.SetAttribute('bus1',cubic_N)
                bus2=Line_item.GetAttribute('bus2').GetAttribute('cterm')
                cubic_A=bus2.CreateObject('StaCubic')
                line5.SetAttribute('bus2',cubic_A)
                cubic_B=bus2.CreateObject('StaCubic')
                line6.SetAttribute('bus2',cubic_B)
                cubic_C=bus2.CreateObject('StaCubic')
                line7.SetAttribute('bus2',cubic_C)
                cubic_N=bus2.CreateObject('StaCubic')
                line8.SetAttribute('bus2',cubic_N)
                towObj = grid.CreateObject('ElmTow', name+'_OHLcoupl2')
                towObj.SetAttribute('pGeo:0', LV_95_ABC)
                towObj.SetAttribute('plines:0', line5)
                towObj.SetAttribute('plines:1', line6)
                towObj.SetAttribute('plines:2', line7)
                towObj.SetAttribute('plines:3', line8)
                line5.GetAttribute('bus1').SetAttribute('it2p1', 0)
                line6.GetAttribute('bus1').SetAttribute('it2p1', 1)
                line7.GetAttribute('bus1').SetAttribute('it2p1', 2)
                line8.GetAttribute('bus1').SetAttribute('it2p1', 3)
                line5.GetAttribute('bus2').SetAttribute('it2p1', 0)
                line6.GetAttribute('bus2').SetAttribute('it2p1', 1)
                line7.GetAttribute('bus2').SetAttribute('it2p1', 2)
                line8.GetAttribute('bus2').SetAttribute('it2p1', 3)
                LV_OH_line_upgrades_idx=LV_OH_line_upgrades_idx+1
                LV_Double_95_ABC_length=LV_Double_95_ABC_length+Line_item.GetAttribute('dline')/1000
            # continue
        else:
            continue
        for type in LV_types:
            typeName = type.GetAttribute('loc_name')
            if typeName == new_type:
                LV_OH_line_upgrades_idx=LV_OH_line_upgrades_idx+1
                Line_item.GetAttribute('c_ptow').SetAttribute('pGeo:0', type)
       
    #### Upgrade underground cables
    elif "cable" in line_name:
        line_type_name = Line_item.GetAttribute('c_ptow').GetAttribute('typ_id').GetAttribute('loc_name')

        if line_type_name == "LV - 240 Al 4C XLPEPVC" or line_type_name == "LV - 240 Cu 4C XLPEPVC" or line_type_name == "LV - 240 Al 3C+NS":
            new_type = "LV - 300 Al 3.5C"
            LV_300_Al_3p5C_length=LV_300_Al_3p5C_length+Line_item.GetAttribute('dline')/1000
        #### Al based cables
        elif line_type_name == "LV - 120 Al 3C+NS":
            new_type = "LV - 240 Al 3C+NS"
            LV_240_Al_3C_NS_length=LV_240_Al_3C_NS_length+Line_item.GetAttribute('dline')/1000
        elif line_type_name == "LV - 120 Al 4C XLPEPVC" or line_type_name == "LV - 150 Al 4C PLYSWS":
            new_type = "LV - 240 Al 4C XLPEPVC"
            LV_240_Al_4C_XLPEPVC_length=LV_240_Al_4C_XLPEPVC_length+Line_item.GetAttribute('dline')/1000
        #### Cu based cables
        elif line_type_name == "LV - 16 Cu 4C PVC":
            new_type = "LV - 25 Cu 4C PLY-HDPE-PVC"
            LV_25_Cu_4C_PLY_HDPE_PVC_length=LV_25_Cu_4C_PLY_HDPE_PVC_length+Line_item.GetAttribute('dline')/1000
        elif line_type_name == "LV - 16 Cu 4C XLPE":
            new_type = "LV - 25 Cu 4C XLPE"
            LV_25_Cu_4C_XLPE_length=LV_25_Cu_4C_XLPE_length+Line_item.GetAttribute('dline')/1000
        elif line_type_name == "LV - 25 Cu 3C+NS PVC" or line_type_name == "LV - 25 Cu 3C+NS XLPE":
            new_type = "LV - 70 Cu 3.5C"
            LV_70_Cu_3p5C_length=LV_70_Cu_3p5C_length+Line_item.GetAttribute('dline')/1000
        elif line_type_name == "LV - 25 Cu 4C PLY-HDPE-PVC" or line_type_name == "LV - 25 Cu 4C XLPE":
            new_type = "LV - 70 Cu 4C HDPE"
            LV_70_Cu_4C_HDPE_length=LV_70_Cu_4C_HDPE_length+Line_item.GetAttribute('dline')/1000
        elif line_type_name == "LV - 70 Cu 3.5C":
            new_type = "LV - 150 Cu 3C+NS"
            LV_150_Cu_3C_NS_length=LV_150_Cu_3C_NS_length+Line_item.GetAttribute('dline')/1000
        elif line_type_name == "LV - 70 Cu 4C HDPE":
            new_type = "LV - 120 Cu 4C XLPEPVC"
            LV_120_Cu_4C_XLPEPVC_length=LV_120_Cu_4C_XLPEPVC_length+Line_item.GetAttribute('dline')/1000
        elif line_type_name == "LV - 120 Cu 4C XLPEPVC" or line_type_name == "LV - 150 Cu 3C+NS":
            new_type = "LV - 240 Cu 4C XLPEPVC"
            LV_240_Cu_4C_XLPEPVC_length=LV_240_Cu_4C_XLPEPVC_length+Line_item.GetAttribute('dline')/1000
        else:
            continue
        for type in LV_types:
            typeName = type.GetAttribute('loc_name')
            if typeName == new_type:
                LV_UG_line_upgrades_idx=LV_UG_line_upgrades_idx+1
                Line_item.GetAttribute('c_ptow').SetAttribute('typ_id', type)

app.PrintPlain(str(LV_OH_spans_scheduled_for_upgrade)+" LV OH spans scheduled for upgrade")
app.PrintPlain(str(LV_OH_line_upgrades_idx)+" LV OH spans upgraded")
app.PrintPlain(str(LV_buried_cable_runs_scheduled_for_upgrade)+" LV buried cable runs scheduled for upgrade")
app.PrintPlain(str(LV_UG_line_upgrades_idx)+" LV buried cable runs upgraded")
app.PrintPlain("Total length of new MOON installed was "+str(MOON_length)+" km")
app.PrintPlain("Total length of new PLUTO installed was "+str(PLUTO_length)+" km")
app.PrintPlain("Total length of new LV_Double_95_ABC installed was "+str(LV_Double_95_ABC_length)+" km")
app.PrintPlain("Total length of new LV_300_Al_3p5C installed was "+str(LV_300_Al_3p5C_length)+" km")
app.PrintPlain("Total length of new LV_240_Al_3C_NS installed was "+str(LV_240_Al_3C_NS_length)+" km")
app.PrintPlain("Total length of new LV_240_Al_4C_XLPEPVC installed was "+str(LV_240_Al_4C_XLPEPVC_length)+" km")
app.PrintPlain("Total length of new LV_25_Cu_4C_PLY_HDPE_PVC installed was "+str(LV_25_Cu_4C_PLY_HDPE_PVC_length)+" km")
app.PrintPlain("Total length of new LV_25_Cu_4C_XLPE installed was "+str(LV_25_Cu_4C_XLPE_length)+" km")
app.PrintPlain("Total length of new LV_70_Cu_3p5C installed was "+str(LV_70_Cu_3p5C_length)+" km")
app.PrintPlain("Total length of new LV_70_Cu_4C_HDPE installed was "+str(LV_70_Cu_4C_HDPE_length)+" km")
app.PrintPlain("Total length of new LV_150_Cu_3C_NS installed was "+str(LV_150_Cu_3C_NS_length)+" km")
app.PrintPlain("Total length of new LV_120_Cu_4C_XLPEPVC installed was "+str(LV_120_Cu_4C_XLPEPVC_length)+" km")
app.PrintPlain("Total length of new LV_240_Cu_4C_XLPEPVC installed was "+str(LV_240_Cu_4C_XLPEPVC_length)+" km")

#### Upgrade the MV lines to the next level
for Line_item in MV_line_set:
    line_type_name = Line_item.GetAttribute('typ_id').GetAttribute('loc_name')

    #### Upgrade overhead lines
    if "FLAT_PIN" in line_type_name:
        conductor_type_name = Line_item.GetAttribute('pCondCir').GetAttribute('loc_name')
        if conductor_type_name == "7/.186 AAC_100°C_11kV_A_SEQ" or conductor_type_name == "HDBC 19/.083_100°C_11kV_A_SEQ" or conductor_type_name == "HDBC 7/.08_100°C_11kV_A_SEQ" or conductor_type_name == "Libra_100°C_11kV_A_SEQ" or conductor_type_name == "Mars_100°C_11kV_A_SEQ":
            new_type = "Moon_100°C_11kV_A_SEQ"
        elif conductor_type_name == "Apple_55°C_11kV_A_SEQ" or conductor_type_name == "Libra_55°C_11kV_A_SEQ":
            new_type = "Moon_55°C_11kV_A_SEQ"
        elif conductor_type_name == "Apple_75°C_11kV_A_SEQ" or conductor_type_name == "Apple_75°C_11kV_B_SEQ" or conductor_type_name == "Banana_75°C_11kV_B_SEQ" or conductor_type_name == "HDBC 19/.083_75°C_11kV_A_SEQ" or conductor_type_name == "HDBC 7/.08_75°C_11kV_A_SEQ" or conductor_type_name == "Libra_75°C_11kV_A_SEQ" or conductor_type_name == "Mars_75°C_11kV_A_SEQ":
            new_type = "Moon_75°C_11kV_A_SEQ"
        else:
            continue
        for type in MV_types:
            typeName = type.GetAttribute('loc_name')
            if typeName == new_type:
                MV_OH_line_upgrades_idx=MV_OH_line_upgrades_idx+1
                Line_item.SetAttribute('pCondCir', type)
    
    #### Upgrade underground cables
    else:
        if line_type_name == "11kVUG.06cuPLYDU":
            new_type = "11kVUG.25cuPLYDU"
        elif line_type_name == "11kVUG.25cuPLYDU" or line_type_name == "11kVUG25cuPLYHDPEDU":
            new_type = "11kVUG95alTRPXDU"
        elif line_type_name == "11kVUG95alTRPXDU" or line_type_name == "11kVUG95alXLDB" or line_type_name == "11kVUG95alXLDU":
            new_type = "11kV120CCT75A"
        elif line_type_name == "11kV120CCT75A":
            new_type = "11kVUG185cuPLYDU"
        elif line_type_name == "11kVUG185cuPLYDU" or line_type_name == "11kVUG185cuPLYHDPDU":
            new_type = "11kVUG240alTRPX90DU"
        elif line_type_name == "11kVUG240alTRPX90DU" or line_type_name == "11kVUG240alTRPX90DU" or line_type_name == "11kVUG240alXL70DU" or line_type_name == "11kVUG240alXL90DU" or line_type_name == "11kVUG240cuPLYDU" or line_type_name == "11kVUG240cuTRPX90DU" or line_type_name == "11kVUG240cuXL90DU":
            new_type = "11kVUG300alPLYDU"
        else:
            continue
        for type in MV_types:
            typeName = type.GetAttribute('loc_name')
            if typeName == new_type:
                MV_UG_line_upgrades_idx=MV_UG_line_upgrades_idx+1
                Line_item.SetAttribute('typ_id', type)

app.PrintPlain(str(len(MV_line_set))+" MV lines scheduled for upgrade")
app.PrintPlain(str(MV_OH_line_upgrades_idx)+" MV OH lines upgraded")
app.PrintPlain(str(MV_UG_line_upgrades_idx)+" MV UG lines upgraded")
#### Upgrade the transformers
all_transformers = app.GetCalcRelevantObjects('ElmTr2')
all_transformer_types = app.GetCalcRelevantObjects('TypTr2')

for item in all_transformers:
    if item.GetAttribute('loc_name') in transformers_to_be_upgraded:
        transformer_type = item.GetAttribute('typ_id').GetAttribute('loc_name')
        if transformer_type == "2-Winding Transformer 25kVA 5 taps":
            new_type = "2-Winding Transformer 30kVA 5 taps"
        elif transformer_type == "2-Winding Transformer 30kVA 5 taps":
            new_type = "2-Winding Transformer 50kVA 5 taps"
        elif transformer_type == "2-Winding Transformer 50kVA 5 taps":
            new_type = "2-Winding Transformer 100kVA 5 taps"
        elif transformer_type == "2-Winding Transformer 100kVA 5 taps":
            new_type = "2-Winding Transformer 200kVA 5 taps"
        elif transformer_type == "2-Winding Transformer 200kVA 5 taps":
            new_type = "2-Winding Transformer 300kVA 5 taps"
        elif transformer_type == "2-Winding Transformer 300kVA 5 taps":
            new_type = "2-Winding Transformer 500kVA 5 taps"
        elif transformer_type == "2-Winding Transformer 25kVA 7 taps":
            new_type = "2-Winding Transformer 63kVA 7 taps"
        elif transformer_type == "2-Winding Transformer 63kVA 7 taps":
            new_type = "2-Winding Transformer 100kVA 7 taps"
        elif transformer_type == "2-Winding Transformer 100kVA 7 taps":
            new_type = "2-Winding Transformer 200kVA 7 taps"
        elif transformer_type == "2-Winding Transformer 200kVA 7 taps":
            new_type = "2-Winding Transformer 315kVA 7 taps"
        elif transformer_type == "2-Winding Transformer 315kVA 7 taps":
            new_type = "2-Winding Transformer 500kVA 7 taps"
        elif transformer_type == "2-Winding Transformer 500kVA 7 taps":
            new_type = "2-Winding Transformer 750kVA 7 taps"
        elif transformer_type == "2-Winding Transformer 750kVA 7 taps":
            new_type = "2-Winding Transformer 1500kVA 7 taps"
        else:
            continue
        for type in all_transformer_types:
                typeName = type.GetAttribute('loc_name')
                if typeName == new_type:
                    Tx_upgrades_idx=Tx_upgrades_idx+1
                    item.SetAttribute('typ_id', type)

app.PrintPlain(str(len(transformers_to_be_upgraded))+" distribution transformers scheduled for upgrade")
app.PrintPlain(str(Tx_upgrades_idx)+" distribution transformers upgraded")
