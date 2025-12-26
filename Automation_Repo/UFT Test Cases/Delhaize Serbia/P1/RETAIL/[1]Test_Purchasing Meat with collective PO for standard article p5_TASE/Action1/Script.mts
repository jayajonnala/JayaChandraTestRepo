
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for standard article p5
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
'.................Test Script Name :Test_Purchasing Meat with collective PO for standard article p5
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Purchasing Meat with collective PO for standard article p5"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------------------MIGO------------------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

''------------------------------Display Article Document Details-------------------------------------------------

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_GODYNPROMAT_DOC, False)
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
Call CheckScreen(DT_SAPTRANSACTIONCODE,DT_MIGO_0750_CHECK_TEXT_OF_TITL)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

'------------------------MIRO-----------------------
Call SetTcode(DT_MIGO_0100_OKCD) 
Call PressEnter()  
Call CheckTCodeScreen(DT_MIGO_0100_OKCD)
Call TakeScreenShot()

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIGO_1000_COMPANY_CODE)
Call TakeScreenShot()
Call PressEnter()  

Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_MIGO_0010_INVOICE_DATE),False)
'Call SetTextbox("","TEXT_BLDAT","",ConvertDate(DT_MIGO_0010_INVOICE_DATE),False)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIGO_0010_REFERENCE,False)
Call PressEnter()
Call TakeScreenShot
Call SetCombo("RM08M-REFERENZBELEGTYP", "Delivery Note")
Call SetTextbox("Incg Doc\. Nmbr","INVFO-INWARDNO_HD","",DT_FB60_0010_REFERENCE,False)
Call SelectCheckBox("INVFO-XMWST","0","ON",False)
Call SetTextBoxNoLabel("RM08M-LFSNR",0,DT_MIGO_6212_RM08MLFSNR,False)
Call PressEnter()
'''Call FindRowNumber("SAPLMR1MTC_MR1M","Purchase order",DT_PO,"DT_ROW_OUTPUT")
Call FindRowNumber("SAPLMR1MTC_MR1M"," Reference Doc.",DT_PO,"DT_ROW_OUTPUT")
'''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","#5",DT_ROW_OUTPUT,"ON",False)



Call TakeScreenShot
Call TakeScreenShot
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIGO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIGO_0010_AMOUNT,False)
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",false)
Call GetStatusBar("item1","DT_MIGO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document no. "&DT_MIGO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")


Call LogOff()
Call FinalStatus ()







'*********************************************End Of Script*********************************************************************

