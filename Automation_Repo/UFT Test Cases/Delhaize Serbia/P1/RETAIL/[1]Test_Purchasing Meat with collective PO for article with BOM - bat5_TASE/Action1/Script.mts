
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for article with BOM - bat5
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

gstrTestCaseName = "Test_Purchasing Meat with collective PO for article with BOM - bat5"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\RETAIL\DT_Purchasing Meat with collective PO for article with BOM - bat5_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INC_REFERENCE",(Cint(DT_INC_REFERENCE)+1))

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''----------------------Tcode MIGO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()   

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0,DT_MIGO_2010_GODYNPROMAT_DOC, False)
Call TakeScreenShot()
Call PressEnter()

Call TakeScreenShot()
'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()

'Click on F1 Document
Call ClickButton("FI Documents",False)
Wait(2)
Call TakeScreenShot()

'''----------------------Tcode MIRO----------------------------

Call SetTcode(DT_MIGO_0750_OKCD) 
Call PressEnter()   

Call SetComboByKey("Transaction",DT_MIGO_6000_TRANSACTION)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_MIGO_0010_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_MIGO_0010_POSTING_DATE),False)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIGO_0010_REFERENCE,False)
Call PressEnter()
Call TakeScreenShot
Call SetCombo("RM08M-REFERENZBELEGTYP", "Delivery Note")
Call SetTextBoxNoLabel("RM08M-LFSNR",0,DT_MIGO_6212_RM08MLFSNR,False)
Call PressEnter()
Call SelectCheckBox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIGO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIGO_0010_AMOUNT,False)
Call PressEnter() 
Call SetTextbox("Incg Doc\. Nmbr","INVFO-INWARDNO_HD","",DT_MIGO_0010_REFERENCE,False)

Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",false)
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_MIGO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document no. "&DT_MIGO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")

Call LogOff()
Call FinalStatus ()


'*********************************************End Of Script*********************************************************************

