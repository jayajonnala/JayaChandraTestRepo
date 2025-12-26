'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_F-44-Clearing vendor 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



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


gstrTestCaseName = "Test_F-44-Clearing vendor"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-F-44----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Call SelectRadioButton("RF05A-XPOS1", "Posting Date", False)
Call SelectRadioButton("RF05A-XPOS1", "Document Number", False)         
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F44_0131_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F44_0131_ACCOUNT	,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_F44_0131_CLEARING_DATE),False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F44_0131_CURRENCY,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("From","RF05A-SEL01","0",DT_F44_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01","1",DT_F44_0731_FROM_OCC1,False)
'Call SetTextbox("From","RF05A-VONDT","",ConvertDate(DT_F44_0131_CLEARING_DATE),False)  
'Call SetTextbox("To","RF05A-SEL02","",DT_F44_0731_FROM,False) 
Call TakeScreenShot
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call VerifyTextBoxContent("Amount entered", "RF05A-BETRG", "", DT_F44_6102_CHECK_TEXT_OF_AMOUNT_ENTERED, False)
Call VerifyTextBoxContent("Not assigned", "RF05A-DIFFB", "", DT_F44_6102_CHECK_TEXT_OF_NOT_ASSIGNED, False)
Call VerifyTextBoxContent("Assigned", "RF05A-NETTO", "", DT_F44_6102_CHECK_TEXT_OF_ASSIGNED, False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)  
Call GetStatusBar("item1", "DT_OP_F44_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",4)
Call VerifyStatusBar(DT_F44_0131_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot
Call SelectMenuBar("Document;Display")
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call VerifyGridCellContent("", 1, "Posting Key","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "Account","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 1, "Description","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Description","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 1, "Amount","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOSTL)
Call VerifyGridCellContent("", 2, "Amount","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)
Call VerifyGridCellContent("", 1, "Currency","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROJK_EXT)
Call VerifyGridCellContent("", 2, "Currency","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROJK_EXT)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
