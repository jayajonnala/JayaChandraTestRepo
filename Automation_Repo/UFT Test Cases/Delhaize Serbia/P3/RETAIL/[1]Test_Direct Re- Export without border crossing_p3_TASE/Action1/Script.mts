
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Direct Re- Export without border crossing_p3
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
'.................Test Script Name :Test_Direct Re- Export without border crossing_p3
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Direct Re- Export without border crossing_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Direct Re- Export without border crossing_p3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode MIRO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(2)

Call SetCombo("RM08M-VORGANG","Invoice")

'Enter the Posting Date
Call SetTextboxNoLabel("INVFO-BLDAT",0,ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)
'Call SetTextbox("Posting Date","INVFO-BUDAT","",DT_MIRO_0010_POSTING_DATE,False)
Wait(2)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
wait(1)
Call PressEnter()
Wait(2)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
'Call TakeScreenShot()
Call PressEnter()
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_MIRO_0010_REFERENCE,False)
Call TakeScreenShot()
Call PressEnter()


'Enter the delivery Note No
Call SetCombo("RM08M-REFERENZBELEGTYP","Purchase Order/Scheduling Agreement")
Call SetTextboxNoLabel("RM08M-EBELN",0,DT_MIRO_6211_RM08MEBELN,False)
Call SetCombo("RM08M-XWARE_BNK","Goods/service items")
Call TakeScreenShot()
Call PressEnter()

'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)
Call TakeScreenShot()

'Get the remaining balance and enter it in Amount Field
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIRO_BALANCE_OUTPUT",False)
Wait(1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT_OCC1,False)
Call PressEnter()
Wait(1)

'Click on Simulate Document
Call ClickButton("Simulate Document   \(Ctrl\+Shift\+F7\)",False) 
Wait(2)
Call TakeScreenShot()

'Click on Post
Call ClickButton("Post   \(Ctrl\+S\)",True) 
Wait(2)

Call ClickButtonIfExist("Save",True)
wait(1)

'Verify message Type
Call VerifyStatusBarMessageType("S")

'Validate If Document No is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
VerifyStatusBar("Document no. "&DT_DOC_NUMBER_OUTPUT&" created")

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


