'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Compensations Process_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 24th March
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Compensations Process_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Compensations Process_p1_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''''Login
'DataRowSet=2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''----------------------Tcode FB70----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_COCODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_COCODE)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_FB70_0510_INVOICE_DATE),"/","."),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace((DT_FB70_0510_POSTING_DATE),"/","."),False)

Call TakeScreenShot()
Call PressEnter()

Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call PressEnter()
Call TakeScreenShot()


Call SelectTab("TS","Basic data",False)
Wait(1)
Call TakeScreenShot()

Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT_1,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_FB70_0510_CALCULATE_TAX,False)
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB70_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FB70_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB70_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB70_0100_TABLECELL_COST_CENTER_0,False)
Call TakeScreenShot()

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Wait(1)
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()
Call TakeScreenShot()


Call VerifyifGuiLabelExists(DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
'Call VerifyifGuiLabelExists(DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
'Call VerifyifGuiLabelExists(DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyifGuiLabelExists(DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyifGuiLabelExists(DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyifGuiLabelExists(DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)


Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_FB70_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_FB70_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyStatusBarMessageType("S")
VerifyStatusBar(DT_FB70_1200_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()


Call LogOff()
Call FinalStatus ()


