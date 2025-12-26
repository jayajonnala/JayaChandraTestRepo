'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_F.03-Clear GL Account  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_F.03-Clear GL Account"
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
'gstrresultFolderPath = ReadTxtFileResult("S:\Resutls_Optimized\AT_F.03-Clear GL Account\GlobalRunTimeResultFolderPath\ResultFolderPath.txt")
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
'
Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("","",DT_MONTH,"Period",False)
Call TakeScreenShot
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","DOCUMENT NUMBER",True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLB03_1105_DOCUMENT_NUMBER,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
'''Call VerifyifGuiLabelExistsByRelativeid("S_LEDR","wnd\[0\]/usr/lbl\[6,8\]")
'Call VerifyGridCellContent("",1,"Cleared/open items symbol","",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
'Call TakeScreenShot


'--------TransactionCode-F-03----------''''

Call SetTcode(DT_FAGLB03_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1) 

Call SelectRadioButton("RF05A-XPOS1","Document Number", False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FAGLB03_0131_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_FAGLB03_0131_ACCOUNT,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_FAGLB03_0131_CURRENCY,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot()
Call SetTextbox("From","RF05A-SEL01","",DT_FAGLB03_0731_FROM,False)
Call TakeScreenShot
Call PressEnter() 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Assigned", "RF05A-AKTIV", "", DT_FAGLB03_6103_CHECK_TEXT_OF_NOT_ASSIGNED, False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_FAGLB03_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
GetRowNo =4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_FAGLB03_0131_CHECK_TEXT_OF_STATUSBAR)

''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_FAGLB03_0131_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR_OCC1,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("","",DT_MONTH,"Period",False)
Call TakeScreenShot
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","DOCUMENT NUMBER",True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLB03_1105_DOCUMENT_NUMBER_OCC1,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call VerifyGridCellContent("",1,"Document Number","",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("",1,"Amount in local currency","",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("",2,"Amount in local currency","",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("",1,"Cleared/open items symbol","",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP_OCC1)
Call VerifyGridCellContent("",1,"Clearing document","",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
'''Call VerifyifGuiLabelExists_ByIndex(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP_OCC1, 0)
'Call VerifyifGuiLabelExists_ByIndex(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR, 0)
'Call VerifyifGuiLabelExists_ByIndex(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB, 0)
'Call VerifyifGuiLabelExists_ByIndex(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB, 0)
'Call ClickButton("Last Column   \(Ctrl\+F4\)",False)
'Call VerifyifGuiLabelExists_ByIndex(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL, 0) 
'Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
