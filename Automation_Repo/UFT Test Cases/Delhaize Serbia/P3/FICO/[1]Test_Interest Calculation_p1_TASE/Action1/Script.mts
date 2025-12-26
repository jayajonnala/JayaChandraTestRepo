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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Reload DataSheet to updates and calculations
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name      : Test_Interest Calculation_p1_TASE
'.................Author : TCS          :Bitan
'................ Creation Date         : 4th May
'.................Modified By           :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Interest Calculation_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Interest Calculation_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'''Login'''
'DataRowSet=2
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''----------------------Tcode FB70----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_COMPANY_CODE1,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_COMPANY_CODE1)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call TakeScreenShot()
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call PressEnter()  
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_FB70_0510_INVOICE_DATE),"/","."),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace((DT_FB70_0510_POSTING_DATE),"/","."),False)
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)
Call PressEnter() 
Call PressEnter() 
Call SelectTab("TS",DT_FB70_1200_BASIC_DATA,False)
Call TakeScreenShot()

Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_FB70_0510_CALCULATE_TAX,False)
Call TakeScreenShot()
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB70_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FB70_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB70_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB70_0100_TABLECELL_COST_CENTER_0,False)
Call TakeScreenShot()

Call SelectTab("TS",DT_FB70_1200_DETAILS,False)
Call TakeScreenShot()

Call ClickButton("Enter",False)
Call SelectTab("TS",DT_FB70_1200_PAYMENT,False)
Call TakeScreenShot()

Call SetTextbox("Pmnt Terms","INVFO-ZTERM","",DT_FB70_0520_PMNT_TERMS,False)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB70_1200_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call TakeScreenShot()


Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call TakeScreenShot()


'----------------------Tcode FB02----------------------------

Call SetTcode(DT_FB70_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FB70_0100_OKCD)
Call TakeScreenShot()

Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB70_0100_FISCAL_YEAR,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB70_0100_COMPANY_CODE,False)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB70_0100_DOCUMENT_NUMBER,False)
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
Call TakeScreenShot()

Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_FB70_0750_CHECK_TEXT_OF_COMPANY_CODE,False)
call VerifyGridCellContent("",1,"Posting Key","",DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"Posting Key","",DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
call VerifyGridCellContent("",3,"Posting Key","",DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
call VerifyGridCellContent("",1,"Account","",DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Account","",DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
call VerifyGridCellContent("",3,"Account","",DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call LogOff()
Call FinalStatus ()

