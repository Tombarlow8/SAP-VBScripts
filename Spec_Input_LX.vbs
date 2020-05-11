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

'Make sure the excel spreadsheet is the active sheet before 
'Accsessing the script through the Macro should ensure this.
Dim objExcel
Dim objSheet
Set objExcel = GetObject(,"Excel.Application") 
'//TODO: is this going to stay the same name for the sheet?
Set objSheet = objExcel.ActiveWorkbook.Worksheets("Latest") 

FG = Trim(CStr(objSheet.Cells(2, 2).Value))
Cycle_counting_Cat = Trim(CStr(objSheet.Cells(111, 5).Value))
Unit_length = Trim(CStr(objSheet.Cells(51, 5).Value))
Unit_width = Trim(CStr(objSheet.Cells(52, 5).Value))
Unit_height = Trim(CStr(objSheet.Cells(53, 5).Value))
Unit_gross_weight = Trim(CStr(objSheet.Cells(54,5).Value))
Unit_vol = Trim(CStr(objSheet.Cells(55,5).Value))
'//TODO: Unit Net weight?
Pak_qty = Trim(CStr(objSheet.Cells(58, 5).Value))
Pak_length = Trim(CStr(objSheet.Cells(59, 5).Value))
Pak_width = Trim(CStr(objSheet.Cells(60, 5).Value))
Pak_height = Trim(CStr(objSheet.Cells(61, 5).Value))
Pak_gross_weight = Trim(CStr(objSheet.Cells(62, 5).Value))
Pak_vol = Trim(CStr(objSheet.Cells(63,5).Value))

Std_qty = Trim(CStr(objSheet.Cells(66, 5).Value))
Std_length = Trim(CStr(objSheet.Cells(67, 5).Value))
Std_width = Trim(CStr(objSheet.Cells(68, 5).Value))
Std_height = Trim(CStr(objSheet.Cells(69, 5).Value))
Std_gross_weight = Trim(CStr(objSheet.Cells(70, 5).Value))
Std_vol = Trim(CStr(objSheet.Cells(71,5).Value))

Lay_qty = Trim(CStr(objSheet.Cells(73, 5).Value))
Lay_length = Trim(CStr(objSheet.Cells(74, 5).Value))
Lay_width = Trim(CStr(objSheet.Cells(75, 5).Value))
Lay_height = Trim(CStr(objSheet.Cells(76, 5).Value))
Lay_gross_weight = Trim(CStr(objSheet.Cells(77, 5).Value))
Lay_vol = Trim(CStr(objSheet.Cells(78,5).Value))

FP_qty = Trim(CStr(objSheet.Cells(80, 5).Value))
FP_length = Trim(CStr(objSheet.Cells(81, 5).Value))
FP_width = Trim(CStr(objSheet.Cells(82, 5).Value))
FP_height = Trim(CStr(objSheet.Cells(83, 5).Value))
FP_gross_weight = Trim(CStr(objSheet.Cells(84, 5).Value))
FP_vol = Trim(CStr(objSheet.Cells(85,5).Value))

Despatchable = Trim(CStr(objSheet.Cells(113,5).Value)) 
Foldable = Trim(CStr(objSheet.Cells(114,5).Value)) 
Non_Conforming = Trim(CStr(objSheet.Cells(115,5).Value))

session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm17"
session.findById("wnd[0]").sendVKey 0

'using the PPD varient to limit the table to UK only measurements (gets rid of South africa and the other measurements)
session.findById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").text = "PPD MARM"
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
objSheet.Cells(38, 2).Value = session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,2]").text

If objSheet.Cells(38, 6).Value = "Pack Size Does Not Match" Then 
    Msgbox "Pack Size cannot be changed From " & session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,2]").text & " to "  & Trim(CStr(objSheet.Cells(8, 8).Value))
    session.findById("wnd[0]/tbar[0]/btn[12]").press
    session.findById("wnd[0]/tbar[0]/btn[12]").press
    session.findById("wnd[0]/tbar[0]/btn[12]").press
    'Stop the script if there is a difference in pack size
    wscript.Quit
End If

'Didn't want to create too many variables so Figured just comments would do.

'Layer Numerator (Units on a layer)
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,0]").text = Trim(CStr(objSheet.Cells(6, 8).Value))
'Standard Numerator (units in a shipper)
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,1]").text = Trim(CStr(objSheet.Cells(7, 8).Value))
'Pack Numerator (units in a pak)
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,2]").text = Trim(CStr(objSheet.Cells(8, 8).Value))
'FP Numerator (units in a FP)
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,3]").text = Trim(CStr(objSheet.Cells(5, 8).Value))
'Unit Numerator (should always be 1)
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD3-VALUE-RIGHT[3,4]").text = "1"
'Layer Height
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,0]").text = Trim(CStr(objSheet.Cells(15, 4).Value))
'Shipper Height
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,1]").text = Trim(CStr(objSheet.Cells(16, 4).Value))
'Pack Height
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,2]").text = Trim(CStr(objSheet.Cells(17, 4).Value))
'FP Height
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,3]").text = Trim(CStr(objSheet.Cells(18, 4).Value))
'Unit Height
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD4-VALUE-RIGHT[4,4]").text = Trim(CStr(objSheet.Cells(14, 4).Value))
'Layer Length
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,0]").text = Trim(CStr(objSheet.Cells(15, 2).Value))
'Standard Length
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,1]").text = Trim(CStr(objSheet.Cells(16, 2).Value))
'Pack Length
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,2]").text = Trim(CStr(objSheet.Cells(17, 2).Value))
'FP Length
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,3]").text = Trim(CStr(objSheet.Cells(18, 2).Value))
'Unit Length
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD5-VALUE-RIGHT[5,4]").text = Trim(CStr(objSheet.Cells(14, 2).Value))
'Layer Volume
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,0]").text = Trim(CStr(objSheet.Cells(15, 6).Value))
'Standard Volume
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,1]").text = Trim(CStr(objSheet.Cells(16, 6).Value))
'Pack Volume
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,2]").text = Trim(CStr(objSheet.Cells(17, 6).Value))
'FP Volume
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,3]").text = Trim(CStr(objSheet.Cells(18, 6).Value))
'unit Volume
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD6-VALUE-RIGHT[6,4]").text = Trim(CStr(objSheet.Cells(14, 6).Value))
'Layer Width
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,0]").text = Trim(CStr(objSheet.Cells(15, 3).Value))
'Standard Width
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,1]").text = Trim(CStr(objSheet.Cells(16, 3).Value))
'Pack Width
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,2]").text = Trim(CStr(objSheet.Cells(17, 3).Value))
'FP width
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,3]").text = Trim(CStr(objSheet.Cells(18, 3).Value))
'Unit Width
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD7-VALUE-RIGHT[7,4]").text = Trim(CStr(objSheet.Cells(14, 3).Value))
'Layer Gross Weight
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,0]").text = Trim(CStr(objSheet.Cells(15, 8).Value))
'Standard Gross Weight
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,1]").text = Trim(CStr(objSheet.Cells(16, 8).Value))
'Pack Gross Weight
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,2]").text = Trim(CStr(objSheet.Cells(17, 8).Value))
'FP Gross Weight
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,3]").text = Trim(CStr(objSheet.Cells(18, 8).Value))
'Unit Gross Weight
session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/txtSTRUC-FIELD8-VALUE-RIGHT[8,4]").text = Trim(CStr(objSheet.Cells(14, 8).Value))

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

'//TODO: do we still need this with LX
'Changing the Net weight of the Unit to be the same as the Gross Weight
session.findById("wnd[0]/tbar[1]/btn[30]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-NTGEW[18,0]").text = Trim(CStr(objSheet.Cells(14, 8).Value))
'session.findById("wnd[0]").sendVKey 0

'Clicking the 'main data' button
session.findById("wnd[0]/tbar[1]/btn[27]").press

'"Enters through" until the warnings have gone
Call RecursiveSAPStatusBarCheck

'Filling in the backscreen 'Warehouse managment 1'
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21").select
'Stock Placement
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").text = Trim(CStr(objSheet.Cells(8, 2).Value))
'Stock Removal
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").text = Trim(CStr(objSheet.Cells(8, 2).Value))
'A or B
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2733/ctxtMLGN-LGBKZ").text = Trim(CStr(objSheet.Cells(9, 2).Value))
'A or B
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2733/ctxtMLGN-BSSKZ").text = Trim(CStr(objSheet.Cells(9, 2).Value))
'Shoppeur Family
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLYFCD:1000/ctxtMLGN-FAMCODE").text = Trim(CStr(objSheet.Cells(10, 2).Value))
'Traceability Family
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLYFCD:1000/ctxtMLGN-FAMTRA").text = Trim(CStr(objSheet.Cells(11, 2).Value))

'//TODO: DO we still need this with the new functionality in YLC01
'Capacity Usage
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2731/txtMLGN-MKAPV").text = Trim(CStr(objSheet.Cells(19, 6).Value))
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2731/ctxtMLGN-BEZME").text = "FP"

'//TODO: dpenedant on warehouse LPD have 'qp' etc
'Filling in the backscreen 'Warehouse managment 2'
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2732/txtMLGN-LHMG1").text = Trim(CStr(objSheet.Cells(5, 2).Value))
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2732/ctxtMLGN-LHME1").text = "un"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2732/ctxtMLGN-LETY1").text = "fp"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2732/txtMLGN-LHMG2").text = Trim(CStr(objSheet.Cells(6, 2).Value))
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2732/ctxtMLGN-LHME2").text = "un"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2732/ctxtMLGN-LETY2").text = "pp"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2732/txtMLGN-LHMG3").text = Trim(CStr(objSheet.Cells(7, 8).Value))
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2732/ctxtMLGN-LHME3").text = "un"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2732/ctxtMLGN-LETY3").text = "std"

'Save
session.findById("wnd[0]").sendVKey 11

'Pauses the program for 1 second otherwise the user locks themselves out of the code when trying to go into MM02
Wscript.Sleep 1000

'//TODO: LX updates procure trade data regardless so don't have to do this anymore
'Re-Entering MM02 to enter the procure trade data as this can only be done in MM02
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = FG
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP33").select
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Looks at the Traceability Family to determine the 'Procure Trade Data' digit
If Trim(CStr(objSheet.Cells(11, 2).Value)) = "5" Then
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP33/ssubTABFRA1:SAPLMGMM:2004/subSUB5:SAPLYMM_BPGV2_2:2016/ctxtYTGRP6-TRACEABILITY").text = "1"
Else
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP33/ssubTABFRA1:SAPLMGMM:2004/subSUB5:SAPLYMM_BPGV2_2:2016/ctxtYTGRP6-TRACEABILITY").text = "0"
End If

'//TODO: if 'Procure Trade Data' digit no longer added in thei swill need to be moved
'Add the cycle counting category in
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = Cycle_counting_Cat
session.findById("wnd[0]/tbar[0]/btn[11]").press

'Save
session.findById("wnd[0]").sendVKey 11



Sub ZUKMM02LX()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZUKMM02LX"
    session.findById("wnd[0]").sendVKey 0
    'Setting the LX flag for material changes
    session.findById("wnd[0]/usr/chkP_ACTIVE").selected = true
    'Setting the flag to take the user to MM02
    session.findById("wnd[0]/usr/chkP_MM02").selected = true
    session.findById("wnd[0]/usr/ctxtP_MATNR").text = FG
    session.findById("wnd[0]/usr/ctxtP_WERKS").text = "LOUK"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]").close
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = FG
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "LOUK"
    session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = "UK3"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
End Sub


Sub YLC01()
'//TODO: change these cell references
    des = Trim(CStr(objSheet.Cells(i,3).Value))
    noncon = Trim(CStr(objSheet.Cells(i,2).Value))
    res = Trim(CStr(objSheet.Cells(i,4).Value))
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nylc01"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = FG
    session.findById("wnd[0]/usr/ctxtRMMG1-LGNUM").text = "uk3"
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    session.findById("wnd[0]/tbar[1]/btn[5]").press
    session.findById("wnd[0]/usr/ctxtYLCD01-YLOCPAR4").text = noncon
    session.findById("wnd[0]/usr/ctxtYLCD01-YLOCPAR5").text = des
    session.findById("wnd[0]/usr/ctxtYLCD01-YLOCPAR6").text = res
    session.findById("wnd[0]/tbar[0]/btn[11]").press
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
