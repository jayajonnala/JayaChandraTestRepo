'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G09P01S10V26 Retire Asset (without customer)  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G09P01S10V26(without)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'''''''--------TransactionCode-ABAON----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call SetTextbox("Company Code","SVALD-VALUE","",DT_ABAON_0300_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABAON_0300_ASSET,False)
Call SetTextbox("Asset","RAIFP2-ANLN2","",DT_ABAON_0300_ASSET_OCC1,False)
Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_ABAON_0200_DOCUMENT_DATE),False)
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_ABAON_0202_ASSET_VALUE_DATE),False)
Call SetTextbox("Manual Revenue","RAIFP2-ERLBT","",DT_ABAON_0410_MANUAL_REVENUE,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_OP_ABAON_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_ABAON_0100_CHECK_TEXT_OF_STATUSBAR)

'''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_ABAON_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)


Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Year-end","","")
Call TakeScreenShot
Call SelectNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Year-end","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Year-end","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Year-end","","")
Call TakeScreenShot
Call DoubleClickGuiGridCell("Transactions", "",2, "Amount posted", False)

''''''--------TransactionCode-FB03----------''''
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot

Call GetGridContentByTitle("", "", "Account", 1, "DT_OP_ABAON_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call GetGridContentByTitle("", "", "Account", 2, "DT_OP_ABAON_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call GetGridContentByTitle("", "", "Account", 3, "DT_OP_ABAON_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR")
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

'''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_ABAON_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)


Call SetTextbox("Account Number","RACCT-LOW","",DT_ABAON_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_ABAON_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_ABAON_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("",1,Month(Date)+1,"Period",False)
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButton("Find",True)
'Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Find","GD_SEARCHSTR","","DOCUMENT NUMBER",True)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","DOCUMENT NUMBER",True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
'Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call ClickButton("Define Filter Values",True)
Call TakeScreenShot
'Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButton("Copy   \(Enter\)",True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ABAON_1105_DOCUMENT_NUMBER_OCC1,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

'Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'Call SelectRowGuiGridbyRowNo("Column Set", 0,2, True)
'Call ClickButton("Add Filter Criterion \(F7\)",True)
'Call TakeScreenShot          
'Call ClickButton("Define Filter Values",True)
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ABAON_1105_DOCUMENT_NUMBER_OCC1,True)
''Call SetTextbox("to","%%DYN001-HIGH","",DT_AB08_1105_TO,True)
'Call ClickButton("Execute   \(Enter\)",True)  
'Call TakeScreenShot
'
Call VerifyGridCellContent("", 1, "Profit Center", "", DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "Amount in local currency", "", DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)

'Call VerifyifGuiLabelExistsByRelativeid(DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR,"wnd\[0\]/usr/lbl\[84,8\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,"wnd\[0\]/usr/lbl\[43,8\]")
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)


Call SetTextbox("Account Number","RACCT-LOW","",DT_ABAON_1000_ACCOUNT_NUMBER_OCC1,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
'Call TakeScreenShot
Call DoubleClickGuiGridCell("",1,Month(Date)+1,"Period",False)
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
'Call ClickButton("Find",True)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Find","GD_SEARCHSTR","","DOCUMENT NUMBER",True)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","DOCUMENT NUMBER",True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
'Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call TakeScreenShot
'Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButton("Copy   \(Enter\)",True)

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ABAON_1105_DOCUMENT_NUMBER_OCC1,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

''
'''Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'''Call SelectRowGuiGridbyRowNo("Column Set", 0,2, True)
'''Call ClickButton("Add Filter Criterion \(F7\)",True)
'''Call TakeScreenShot          
'''Call ClickButton("Define Filter Values",True)
'''Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ABAON_1105_DOCUMENT_NUMBER_OCC1,True)
''''Call SetTextbox("to","%%DYN001-HIGH","",DT_AB08_1105_TO,True)
'''Call ClickButton("Execute   \(Enter\)",True)  
'''Call TakeScreenShot
'''
Call VerifyGridCellContent("", 1, "Profit Center", "", DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR_OCC1)
Call VerifyGridCellContent("", 1, "Amount in local currency", "", DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call TakeScreenShot
'Need this code for the Guilabel verifications
'Call VerifyifGuiLabelExistsByRelativeid(DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR_OCC1,"wnd\[0\]/usr/lbl\[84,8\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1,"wnd\[0\]/usr/lbl\[43,8\]")
'Call TakeScreenShot

Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

Call SetTextbox("Account Number","RACCT-LOW","",DT_ABAON_1000_ACCOUNT_NUMBER_OCC2,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("",1,Month(Date)+1,"Period",False)
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
'Call ClickButton("Find",True)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
'Call SetTextbox("Find","GD_SEARCHSTR","","DOCUMENT NUMBER",True)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","DOCUMENT NUMBER",True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
'Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call TakeScreenShot
'Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButton("Copy   \(Enter\)",True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ABAON_1105_DOCUMENT_NUMBER_OCC1,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
'
''''Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
''''Call SelectRowGuiGridbyRowNo("Column Set", 0,2, True)
''''Call ClickButton("Add Filter Criterion \(F7\)",True)
''''Call TakeScreenShot          
''''Call ClickButton("Define Filter Values",True)
''''Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ABAON_1105_DOCUMENT_NUMBER_OCC1,True)
'''''Call SetTextbox("to","%%DYN001-HIGH","",DT_AB08_1105_TO,True)
''''Call ClickButton("Execute   \(Enter\)",True)  
''''Call TakeScreenShot
'''
Call VerifyGridCellContent("", 1, "Profit Center", "", DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR_OCC2)
Call VerifyGridCellContent("", 1, "Amount in local currency", "", DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC2)
Call TakeScreenShot

'Call VerifyifGuiLabelExistsByRelativeid(DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR_OCC2,"wnd\[0\]/usr/lbl\[84,8\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_ABAON_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC2,"wnd\[0\]/usr/lbl\[43,8\]")
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

