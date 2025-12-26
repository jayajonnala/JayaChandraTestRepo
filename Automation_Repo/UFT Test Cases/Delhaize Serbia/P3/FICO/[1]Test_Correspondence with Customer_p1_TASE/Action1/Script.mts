'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Correspondence with Customer_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 08th March
'.................Modified By :
'.................Modified Date/Details :
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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Correspondence with Customer_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Correspondence with Customer_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
''Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''----------------------Tcode FB70----------------------------
''Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
''Capture the screenshot
Call TakeScreenShot()
''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_COMPANY_CODE,True)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_COMPANY_CODE)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call PressEnter()
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_FB70_0510_INVOICE_DATE),"/","."),False)
Call PressEnter()
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)
Call PressEnter()
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
'select Basic data tab
Call SelectTab("TS","Basic data",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT,False)
Call PressEnter()
Call PressEnter()
''Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_FB70_0510_CALCULATE_TAX,False)
''Capture the screenshot
Call TakeScreenShot()
''set table data
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB70_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FB70_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB70_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB70_0100_TABLECELL_COST_CENTER_0,False)
''Capture the screenshot
Call TakeScreenShot()
''select Basic data tab
Call SelectTab("TS","Details",False)
''Capture the screenshot
Call TakeScreenShot()
Wait(2)
''Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
'Click on Post Buton
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If document is posted and get the status bar nummber
Call GetStatusBar("item1","DT_FB70_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB70_1200_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

