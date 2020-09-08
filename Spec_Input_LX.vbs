If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").resizeWorkingPane 100,39,false

'Github link: https://github.com/Tombarlow8/SAP-VBScripts/blob/master/Spec_Input_LX.vbs 
'Make sure the excel spreadsheet is the active sheet before 
'Accsessing the script through the Macro should ensure this.
Dim objExcel
Dim objSheet
Set objExcel = GetObject(,"Excel.Application") 
'//TODO: is this going to stay the same name for the sheet?
Set objSheet = objExcel.ActiveWorkbook.Worksheets("Latest") 

FG = Trim(CStr(objSheet.Cells(2, 2).Value))
Cycle_counting_Cat = Trim(CStr(objSheet.Cells(111, 5).Value))
Warehouse_number = Trim(CStr(objSheet.Cells(4, 2).Value))
Plant = Trim(CStr(objSheet.Cells(118, 5).Value))

Unit_ean = Trim(CStr(objSheet.Cells(50, 5).Value))
Unit_length = Trim(CStr(objSheet.Cells(51, 5).Value))
Unit_width = Trim(CStr(objSheet.Cells(52, 5).Value))
Unit_height = Trim(CStr(objSheet.Cells(53, 5).Value))
Unit_gross_weight = Trim(CStr(objSheet.Cells(54,5).Value))
Unit_vol = Trim(CStr(objSheet.Cells(55,5).Value))
Unit_Net_weight = Trim(CStr(objSheet.Cells(56,5).Value))

Pak_ean = Trim(CStr(objSheet.Cells(57, 5).Value))
Pak_qty = Trim(CStr(objSheet.Cells(58, 5).Value))
Pak_length = Trim(CStr(objSheet.Cells(59, 5).Value))
Pak_width = Trim(CStr(objSheet.Cells(60, 5).Value))
Pak_height = Trim(CStr(objSheet.Cells(61, 5).Value))
Pak_gross_weight = Trim(CStr(objSheet.Cells(62, 5).Value))
Pak_vol = Trim(CStr(objSheet.Cells(63,5).Value))

Std_ean = Trim(CStr(objSheet.Cells(65, 5).Value))
Std_qty = Trim(CStr(objSheet.Cells(66, 5).Value))
Std_length = Trim(CStr(objSheet.Cells(67, 5).Value))
Std_width = Trim(CStr(objSheet.Cells(68, 5).Value))
Std_height = Trim(CStr(objSheet.Cells(69, 5).Value))
Std_gross_weight = Trim(CStr(objSheet.Cells(70, 5).Value))
Std_vol = Trim(CStr(objSheet.Cells(71,5).Value))

Lay_ean = Trim(CStr(objSheet.Cells(72, 5).Value))
Lay_qty = Trim(CStr(objSheet.Cells(73, 5).Value))
Lay_length = Trim(CStr(objSheet.Cells(74, 5).Value))
Lay_width = Trim(CStr(objSheet.Cells(75, 5).Value))
Lay_height = Trim(CStr(objSheet.Cells(76, 5).Value))
Lay_gross_weight = Trim(CStr(objSheet.Cells(77, 5).Value))
Lay_vol = Trim(CStr(objSheet.Cells(78,5).Value))

FP_ean = Trim(CStr(objSheet.Cells(79, 5).Value))
FP_qty = Trim(CStr(objSheet.Cells(80, 5).Value))
FP_length = Trim(CStr(objSheet.Cells(81, 5).Value))
FP_width = Trim(CStr(objSheet.Cells(82, 5).Value))
FP_height = Trim(CStr(objSheet.Cells(83, 5).Value))
FP_gross_weight = Trim(CStr(objSheet.Cells(84, 5).Value))
FP_vol = Trim(CStr(objSheet.Cells(85,5).Value))

Support_Code = Trim(CStr(objSheet.Cells(87,5).Value))
Support_Height = Trim(CStr(objSheet.Cells(88,5).Value))
Boxes_per_Layer = Trim(CStr(objSheet.Cells(89,5).Value))
Layers_per_Pallet = Trim(CStr(objSheet.Cells(90,5).Value))

Shopeur_family = Trim(CStr(objSheet.Cells(94,5).Value))
Trace_family = Trim(CStr(objSheet.Cells(95,5).Value))
Stock_Removal_ind = Trim(CStr(objSheet.Cells(96,5).Value))
Stock_placement_ind = Trim(CStr(objSheet.Cells(97,5).Value))
Storage_section_ind = Trim(CStr(objSheet.Cells(98,5).Value))
Special_movement_ind = Trim(CStr(objSheet.Cells(99,5).Value))

Despatchable = Trim(CStr(objSheet.Cells(113,5).Value)) 
Non_Conforming = Trim(CStr(objSheet.Cells(115,5).Value)) 
Restrict = Trim(CStr(objSheet.Cells(114,5).Value))

Do_not_use_layer = Trim(CStr(objSheet.Cells(116,5).Value))
Do_not_use_Pallet = Trim(CStr(objSheet.Cells(117,5).Value))


Call ZUKMM02LX
Call MM17
Call MM02
Call YLC01


Sub MM02()        
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = FG
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]").sendVKey 0

    'If an error pops up in the status bar while in MM02 it will create the code in MM01 instead
    If session.FindById("wnd[0]/sbar").Text = "Select at least one view" Then
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = FG
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[1]").sendVKey 0
        session.findById("wnd[1]").sendVKey 0
    End If

    
    'Changing the Net weight of the Unit as it can only be done in MM01/2
    session.findById("wnd[0]/tbar[1]/btn[30]").press
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02").select
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-NTGEW[18,0]").text = Unit_Net_weight
    'session.findById("wnd[0]").sendVKey 0
    
    'Clicking the 'main data' button
    session.findById("wnd[0]/tbar[1]/btn[27]").press
    
    '"Enters through" until the warnings have gone needs an initial enter to start
    'session.findById("wnd[1]").sendVKey 0
    Call RecursiveSAPStatusBarCheck

    'Filling in the backscreen 'Warehouse managment 1'
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21").select
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").text = Stock_placement_ind
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").text = Stock_Removal_ind
    '//TODO: check if these 2 are the correct way round
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2733/ctxtMLGN-LGBKZ").text = Storage_section_ind
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2733/ctxtMLGN-BSSKZ").text = Special_movement_ind

    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLYFCD:1000/ctxtMLGN-FAMCODE").text = Shopeur_family
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLYFCD:1000/ctxtMLGN-FAMTRA").text = Trace_family

    'Add the cycle counting category in
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19").select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = Cycle_counting_Cat
    

    'Maintain the manutention & label data
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP34").select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP34/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLYMM_BPGV2_2:2002/ctxtYTGRP11-TYSUP").text = Support_Code
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP34/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLYMM_BPGV2_2:2002/txtYTGRP11-LAYNB").text = Layers_per_Pallet
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP34/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLYMM_BPGV2_2:2002/txtYTGRP11-YNBPROD_LAY").text = Boxes_per_Layer
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP34/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLYMM_BPGV2_2:2002/txtYTGRP11-HOEHE_MAN").text = Support_Height    

    'Save
    session.findById("wnd[0]/tbar[0]/btn[11]").press

End sub


Sub MM17()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm17"
    session.findById("wnd[0]").sendVKey 0

    'using the PPD varient to limit the table to UK only measurements (gets rid of South africa and the other measurements)
    session.findById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").text = "PPD UoM"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[0,24]").text = FG


    '//TODO: maybe get rid of noone checks anyway and can't see any reason it would be different form spreadsheet (Do the check in the actual spreadsheet?)
    answer = msgbox ("Check FG code before clicking Yes", vbYesNo)
    'If the FG is incorrect for any reason the script should stop (6 = "Yes")
    If answer <> 6 Then
        session.findById("wnd[0]/tbar[0]/btn[12]").press
        session.findById("wnd[0]/tbar[0]/btn[12]").press
        wscript.Quit
    End If
    
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    '//TODO: Alisatirs spreadsheet could include this check. With LX 
    'Grabs the pack size from SAP as this should not be changed initially by Stock Control
    objSheet.Cells(2, 7).Value = session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,2]").text
    msgbox objSheet.Cells(2, 7).Value
    If objSheet.Cells(2, 7).Value = "Pack Size Does Not Match" Then 
        Msgbox "Pack Size cannot be changed From " & session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,2]").text & " to "  & Trim(CStr(objSheet.Cells(8, 8).Value))
        session.findById("wnd[0]/tbar[0]/btn[12]").press
        session.findById("wnd[0]/tbar[0]/btn[12]").press
        session.findById("wnd[0]/tbar[0]/btn[12]").press
        'Stop the script if there is a difference in pack size
        wscript.Quit
    End If

    'Filling in the UoM in MM17
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,0]").text = Lay_qty
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,1]").text = Std_qty
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,2]").text = Pak_qty
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,3]").text = FP_qty
    'Unit Numerator (should always be 1)
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,4]").text = "1"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,0]").text = Lay_height
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,1]").text = Std_height
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,2]").text = Pak_height
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,3]").text = FP_height
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,4]").text = Unit_height
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,0]").text = Lay_length
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,1]").text = Std_length
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,2]").text = Pak_length
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,3]").text = FP_length
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,4]").text = Unit_length
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,0]").text = Lay_vol
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,1]").text = Std_vol
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,2]").text = Pak_vol
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,3]").text = FP_vol
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,4]").text = Unit_vol
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,0]").text = Lay_width
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,1]").text = Std_width
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,2]").text = Pak_width
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,3]").text = FP_width
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,4]").text = Unit_width
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,0]").text = Lay_gross_weight
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,1]").text = Std_gross_weight
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,2]").text = Pak_gross_weight
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,3]").text = FP_gross_weight
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,4]").text = Unit_gross_weight
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD12-VALUE-LEFT[12,0]").text = Lay_ean
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD12-VALUE-LEFT[12,1]").text = Std_ean
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD12-VALUE-LEFT[12,2]").text = Pak_ean
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD12-VALUE-LEFT[12,3]").text = FP_ean
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD12-VALUE-LEFT[12,4]").text = Unit_ean
    'These units of measure should be always the same therefore hardcoded
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD9-VALUE-LEFT[9,0]").text = "MM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD9-VALUE-LEFT[9,1]").text = "MM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD9-VALUE-LEFT[9,2]").text = "MM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD9-VALUE-LEFT[9,3]").text = "MM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD9-VALUE-LEFT[9,4]").text = "MM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD10-VALUE-LEFT[10,0]").text = "CCM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD10-VALUE-LEFT[10,1]").text = "CCM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD10-VALUE-LEFT[10,2]").text = "CCM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD10-VALUE-LEFT[10,3]").text = "CCM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD10-VALUE-LEFT[10,4]").text = "CCM"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD11-VALUE-LEFT[11,0]").text = "G"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD11-VALUE-LEFT[11,1]").text = "G"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD11-VALUE-LEFT[11,2]").text = "G"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD11-VALUE-LEFT[11,3]").text = "G"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD11-VALUE-LEFT[11,4]").text = "G"

    session.findById("wnd[0]/tbar[0]/btn[11]").press

    '//TODO: add error handling here there is a success/error message we can pull out
    Call MM17StatusCheck

End Sub


Sub ZUKMM02LX()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZUKMM02LX"
    session.findById("wnd[0]").sendVKey 0
    'Setting the LX flag for material changes
    session.findById("wnd[0]/usr/chkP_ACTIVE").selected = true
    'Setting the flag to take the user to MM02
    session.findById("wnd[0]/usr/chkP_MM02").selected = false
    session.findById("wnd[0]/usr/ctxtP_MATNR").text = FG
    session.findById("wnd[0]/usr/ctxtP_WERKS").text = Plant
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]").sendVKey 0
    'if the LX tick is alrwady in no message box pop up is needed to close
    on error resume next
    session.findById("wnd[1]").close
End Sub


Sub YLC01()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nYLC01"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = FG
    session.findById("wnd[0]/usr/ctxtRMMG1-LGNUM").text = Warehouse_number
    session.findById("wnd[0]/tbar[1]/btn[5]").press
    ' IF status bar...
    If session.FindById("wnd[0]/sbar").Text = "This complete material key already exist in table YLCD01" Then
        session.findById("wnd[0]/sbar").doubleClick
        session.findById("wnd[0]/shellcont").close
        session.findById("wnd[0]/tbar[1]/btn[6]").press
        session.findById("wnd[0]/tbar[1]/btn[14]").press
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/tbar[1]/btn[5]").press
    End If

    If Do_not_use_Pallet = "X" Then
        session.findById("wnd[0]/usr/chkYLCD01-YP_NOTUSE").selected = true
    Else
        session.findById("wnd[0]/usr/chkYLCD01-YP_NOTUSE").selected = false
    End If

    If Do_not_use_layer = "X" Then
        session.findById("wnd[0]/usr/chkYLCD01-YL_NOTUSE").selected = true
    Else
        session.findById("wnd[0]/usr/chkYLCD01-YL_NOTUSE").selected = false
    End If 
    session.findById("wnd[0]/usr/ctxtYLCD01-YLOCPAR4").text = Non_Conforming
    session.findById("wnd[0]/usr/ctxtYLCD01-YLOCPAR5").text = Despatchable
    session.findById("wnd[0]/usr/ctxtYLCD01-YLOCPAR6").text = Restrict
    'Fragile?
    'session.findById("wnd[0]/usr/txtYLCD01-YLOCPAR7").text = "X"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    session.findById("wnd[1]/usr/btnBUTTON_1").press
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
End Sub


'For FG Codes which have other countries data in MM02/1 this Sub checks the status
'bar for warnings related to net weight/gross weight incompatibility  and 'Enters through' until the 
'warnings have gone. the Recursion is used as the amount of warnings is ambiguous.  
Sub RecursiveSAPStatusBarCheck()
    session.findById("wnd[0]").sendVKey 0
    If session.FindById("wnd[0]/sbar").Text = "The net weight is greater than the gross weight" Then 'FYI this is case sensative 		
        Call RecursiveSAPStatusBarCheck 
    Else
        Exit Sub    
    End If
End Sub



Sub MM17StatusCheck()
    value =  session.findById("wnd[0]/usr/tblSAPLMASSMSGLISTTC_MSG/lblLIGHT[0,0]").IconName
    if value = "S_TL_G" Then
        Exit Sub 
    elseif value = "S_TL_R" Then
        msgbox = "There is an error in the inputs please see details"
        wscript.Quit
    elseif value = "S_TL_Y" Then
        msgbox = "There is an warning error in the inputs please see details"
        wscript.Quit
    End if
End Sub
