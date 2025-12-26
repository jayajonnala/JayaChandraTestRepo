'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G09P01S10V27 Retire Asset by scrapping 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G09P01S10V27 Retire Ascrapping"
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

''''Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-ABAVN----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","SVALD-VALUE","",DT_ABAVN_0300_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABAVN_0300_ASSET,False)  
Call SetTextbox("Asset","RAIFP2-ANLN2","",DT_ABAVN_0300_ASSET_OCC1,False)  
Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_ABAVN_0200_DOCUMENT_DATE),False)
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_ABAVN_0202_ASSET_VALUE_DATE),False)
Call SetTextbox("Posting Date","RAIFP1-BUDAT","",ConvertDate(DT_ABAVN_0201_POSTING_DATE),False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call SetTextbox("Text","RAIFP2-SGTXT","",DT_ABAVN_0206_TEXT,False)  
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_OP_DOC")

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_ABAVN_0100_CHECK_TEXT_OF_STATUSBAR)

''''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_ABAVN_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","","")
Call TakeScreenShot
Call DoubleClickGuiGridCell("Transactions", "",2, "Amount posted", False)

''''''--------TransactionCode-FB03----------''''

Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot


Call VerifyTextBoxContent("Document Number", "BKPF-BELNR", "", DT_ABAVN_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call GetGridContentByTitle("", "", "Account", 2, "DT_OP_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR")

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call TakeScreenShot

''''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_ABAVN_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)


Call SetTextbox("Account Number","RACCT-LOW","",DT_ABAVN_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_ABAVN_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_ABAVN_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("",1,Month(Date)+1,"Period",False)
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
' ClickButton(tooltipOrButtonName, blnIsItPopup)
Call ClickButton("Find",True)


Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ABAVN_1105_DOCUMENT_NUMBER,True)
Call ClickButton("Execute   \(Enter\)",True) 


'Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'Call ClickButtonToolBar("&FIND",0)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
'Call SelectCheckbox("GS_SEARCH-EXACT_WORD",0, "ON", True)
'Call ClickButton("OK   \(Enter\)",True)
'Call ClickButton("Cancel   \(F12\)",True)
'Call ClickButton("Add Filter Criterion \(F7\)",True)
'Call TakeScreenShot          
'Call ClickButton("Define Filter Values",True)
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ABAVN_1105_DOCUMENT_NUMBER,True)
'Call ClickButton("Execute   \(Enter\)",True) 


Call TakeScreenShot

Call VerifyifGuiLabelExistsByRelativeid(DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SEGMENT,"wnd\[0\]/usr/lbl\[95,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR,"wnd\[0\]/usr/lbl\[84,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,"wnd\[0\]/usr/lbl\[43,8\]")

'Call VerifyGridCellContent("", 1, "Segment", "0", DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SEGMENT)
'Call VerifyGridCellContent("", 1, "Profit Center", "0", DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
'Call VerifyGridCellContent("", 1, "Amount in local currency", "0", DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call TakeScreenShot

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

