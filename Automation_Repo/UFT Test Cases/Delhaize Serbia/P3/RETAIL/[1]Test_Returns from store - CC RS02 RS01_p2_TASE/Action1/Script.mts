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
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Returns from store - CC RS02 RS01_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 10th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Returns from store - CC RS02 RS01_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Returns from store - CC RS02 RS01_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''''Reload DataSheet to updates and calculations
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_REF_INC",(Cint(DT_REF_INC)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_6150_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOHEAD",DT_MIGO_0100_DOC_INFO,False)
'Capture the screenshot
Call TakeScreenShot()

'Capture the screenshot
'Call TakeScreenShot()

'___________________following block is as per log, but this flow is obslute. If script fails following can be used___________


Call ClickButton("FI Documents",False)
Wait(1)
Call SelectRowGuiGridbyRowNo("Documents in Accounting","",DT_MIGO_0200_GRIDCELL_0_DOC_NUMBER,true)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Display Document   \(F2\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
'__________________________________________________________end block_________________________________________________________
'
'-------------------------------------------------MIRO-----------------------------------------
Call SetTcode(DT_MIGO_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_0750_OKCD)

Call SetComboByKey("RM08M-VORGANG",DT_MIGO_6000_TRANSACTION)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

'Call SetTextbox("Invoice date","INVFO-BLDAT","",DT_MIGO_0010_INVOICE_DATE,False)
Call SetTextboxNoLabel("INVFO-BLDAT","",Replace(DT_MIGO_0010_INVOICE_DATE,"/","."),False)
Call PressEnter() 
wait(1)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIGO_0010_REFERENCE,False)
Call PressEnter()
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_MIGO_0010_REFERENCE,False)
Call PressEnter()
Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace(DT_MIGO_0010_POSTING_DATE,"/","."),False)
'Enter the delivery Note No
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_MIGO_6020_RM08MREFERENZBELEGTYP)

'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIGO_6212_RM08MLFSNR,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("HEADER",DT_MIGO_6005_DETAILS,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("INVFO-BLART",DT_MIGO_0150_DOC_TYPE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()
VerifyStatusBar(DT_MIGO_6000_CHECK_TEXT_OF_STATUSBAR)
VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call SelectTab("HEADER",DT_MIGO_6005_BASIC_DATA,False)
'Capture the screenshot
Call TakeScreenShot()

'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_MIGO_0010_CALCULATE_TAX,False)
Call TakeScreenShot()

'Get the remaining balance and enter it in Amount Field
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIGO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIGO_0010_AMOUNT,False)
Call PressEnter()
Wait(1)
Call TakeScreenShot()

'Click on Post Buton
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_MIGO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_MIGO_6000_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call VerifyStatusBarMessageType("S")

'-----------------------------------------------MIR4-----------------------------------------------------
Call SetTcode(DT_MIGO_6000_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_MIGO_6000_OKCD)
Call TakeScreenShot()
'Display the Invoice Details
Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIGO_6150_INVOICE_DOCUMENT_NO,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIGO_6150_FISCAL_YEAR,False)
Call TakeScreenShot()
Call PressEnter()
wait(1)
Call TakeScreenShot()
'Click on Follow On Document
Call ClickButtonIfExist("Follow-On Documents \.\.\.   \(F8\)",False)
wait(1)
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


