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

''Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name      : Test_PreReq_Manage AR Bill of Exchange_p3_TASE
'.................Author : TCS          : Bitan
'................ Creation Date         : 20th April
'.................Modified By           :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PreReq_Manage AR Bill of Exchange_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_PreReq_Manage AR Bill of Exchange_p3_TASE.xls"

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''----------------------Tcode FB70----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()
	
'this step is not in the log but it is a mandate in screen
'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_COMPANY_CODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_COMPANY_CODE)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(1)
Call TakeScreenShot()

Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDateFormat(DT_FB70_0510_INVOICE_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDateFormat(DT_FB70_0510_POSTING_DATE),False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT,False)
'Call SetTextbox("Amount","INVFO-WAERS","",DT_FB70_0510_AMOUNT_OCC1,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_FB70_0510_CALCULATE_TAX,False)

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(1)

Call SetComboByKey("INVFO-BLART",DT_DOC_TYPE)
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call PressEnter() 
Call TakeScreenShot()

Call SelectRowGuiTableByRow("SAPLFSKBTABLE",1,False)
Call SendKey("{TAB}")
Call TakeScreenShot()
Wait(2)
Call SendKey("{F4}")
Wait(2)
'Call SetTextbox("Language Key","G_SELFLD_TAB-LOW","","",True)
Call SetTextbox("G/L account","G_SELFLD_TAB-LOW","",DT_FB70_0100_TABLECELL_GL_ACCT_0,True)
Call TakeScreenShot()
Call PressEnter() 
Call ClickButton("Find   \(Ctrl\+F\)",True)

Call SetTextbox("Find","RSYSF-STRING","",DT_FIND,True)
Call TakeScreenShot()
Call ClickButton("Find   \(Enter\)",True)


Call TakeScreenShot()
Call ClickLabel(DT_FIND,"",True)

Call TakeScreenShot()
Call ClickButton("Copy   \(Enter\)",True)
'Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB70_0100_TABLECELL_GL_ACCT_0,False)

Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FB70_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB70_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB70_0100_TABLECELL_COST_CENTER_0,False)
Call PressEnter()
Call TakeScreenShot()

Call SelectTab("TS","Details",False)
Call TakeScreenShot()

Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call TakeScreenShot()

Call SelectTab("TS","Payment",False)
Call TakeScreenShot()

Call SelectTab("TS","Tax",False)
Call TakeScreenShot()

Call SelectTab("TS","Notes",False)
Call TakeScreenShot()

Call SetTextArea(DT_FB70_0540_TEXTEDIT_SHELL) 
Call TakeScreenShot()

Call SelectTab("TS","Basic data",False)
Call TakeScreenShot()

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Wait(1)
Call TakeScreenShot()
Call PressEnter()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_FB70_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB70_1200_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()


Call LogOff()
Call FinalStatus ()

