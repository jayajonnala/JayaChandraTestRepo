
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Regular purchasing in SW31 dry goods_p3
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
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Regular purchasing in SW31 dry goods_p3
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Regular purchasing in SW31 dry goods_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Regular purchasing in SW31 dry goods_p3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode MIRO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(2)

'Enter the Posting Date
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_MIRO_0010_POSTING_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call SetTextbox("Incg Doc\. Nmbr","INVFO-INWARDNO_HD","",DT_MIRO_0010_REFERENCE,False)
Call TakeScreenShot()
Call PressEnter()


'Enter the delivery Note No
Call SetCombo("RM08M-REFERENZBELEGTYP","Delivery Note")
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIRO_6212_RM08MLFSNR,False)
Call TakeScreenShot()
Call PressEnter()

'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,"ON",False)

'Get the remaining balance and enter it in Amount Field
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIRO_BALANCE_OUTPUT",False)
Wait(1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT,False)
Call PressEnter()
Wait(1)

'Verify Balance
Call VerifyTextBoxContent("Balance","RM08M-DIFFERENZ",0,DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OCC1,False)

'Select Row
Call SelectRowGuiTable("SAPLMR1MTC_MR1M","Item",1,False)

'Click on Post Buton
Call SelectMenuBar("Invoice Document;Post")
Call ClickButtonIfExist("Save",True)
wait(2)

'Validate If Purchase order is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
VerifyStatusBar("Document no. "&DT_DOC_NUMBER_OUTPUT&" created")

'----------------------Tcode MIR4----------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIRO_6000_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_MIRO_6000_OKCD)

'Display the Invoice Details
Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIRO_6150_INVOICE_DOCUMENT_NO,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIRO_6150_FISCAL_YEAR,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Display Document   \(F2\)",False)
wait(2)

'Click on Follow On Document
Call ClickButtonIfExist("Follow-On Documents \.\.\.   \(F8\)",False)
wait(2)

'Verify the screen title
Call CheckScreen(DT_MIRO_6000_OKCD,DT_MIRO_0750_CHECK_TEXT_OF_TITL)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)
Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

