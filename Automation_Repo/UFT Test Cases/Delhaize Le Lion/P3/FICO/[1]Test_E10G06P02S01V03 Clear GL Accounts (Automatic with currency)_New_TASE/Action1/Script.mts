'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E10G06P02S01V03 Clear GL Accounts (Automatic with currency)  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "TC3_Test_Clear(Auto curr)"
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

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-F-03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SelectRadioButton("RF05A-XPOS1","Document Number", False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F03_0131_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F03_0131_ACCOUNT,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F03_0131_CURRENCY,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F03_0131_PERIOD,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_F03_0131_CLEARING_DATE),False)
Call TakeScreenShot
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call SetTextbox("From","RF05A-SEL01","",DT_F03_0731_FROM,False)
Call SetTextbox("To","RF05A-SEL02","",DT_F03_0731_TO,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)   
Call TakeScreenShot()
Call ClickButtonIfExist("Select All",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Deactivate Items",False)
Call TakeScreenShot()
''Call VerifyTextBoxContent("Not assigned", "RF05A-DIFFB", "", DT_F03_6103_CHECK_TEXT_OF_NOT_ASSIGNED, False)
'''Call VerifyTextBoxContent("Not assigned", "RF05A-DIFFB", "", "0,00", False)
Call ClickButtonIfExist("Select All",False)
Call ClickButtonIfExist("Activate Items",False)
Call VerifyTextBoxContent("Not assigned", "RF05A-DIFFB", "", "0,00", False)
Call ClickButtonIfExist("Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetTextStatusBar("DT_OP_F03_0131_CHECK_TEXT_OF_STATUSBAR")
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)

''--------TransactionCode-FB03----------''''

Call SetTcode(DT_F03_0131_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)


'Call SetTextbox("Document Number","RF05L-BELNR","",DT_F03_0100_DOCUMENT_NUMBER,False)
'Call SetTextbox("Doc. Number","RF05L-BELNR","",DT_F03_0100_DOCUMENT_NUMBER,False)
Call SetTextboxNoLabel("RF05L-BELNR","",DT_F03_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F03_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F03_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiGridbyRowNo("",0,2,False)
Call ClickButtonToolBar("&DETAIL",0)


Call FindRowNumber("", "Group description", "Amt.in loc.cur.", "DT_ROW_NUMBER")
Call GetGridContent("",0, "Contents",DT_ROW_NUMBER,"Column","Company code", "DT_OP_F03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR")
Call ClickButton("Close window   \(Enter\)",True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)


Call ClickContextButtonToolBar("&MB_VARIANT",1)
Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_SEARCH_TERM,True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("",1,"Profit Center",0, DT_F03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",1,"Account",0, DT_ACCNT_1)
'Call VerifyGridCellContent("",2,"Cost Center",0, DT_F03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)
'Call VerifyGridCellContent("",2,"Account",0, DT_ACCNT)
Call VerifyGridCellContent("",2,"Profit Center",0, DT_F03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
'Call VerifyGridCellContent("",2,"Description",0, DT_F03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)
Call TakeScreenShot

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

