
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Manual Posting of GL Document  ledger specific_p1
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



'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manual Posting of GL Document  ledger specific_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manual Posting of GL Document ledger specific_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Manual Posting of GL Document  ledger specific_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode FB50L----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 

'Enter the conpany code
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_CONTROLLING_AREA)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(2)

'Enter the details
Call SetTextbox("Document Date","ACGL_HEAD-BLDAT","",DT_FB50L_1010_DOCUMENT_DATE,False)   
Call SetTextbox("Ledger Grp","ACGL_HEAD-LDGRP","",DT_FB50L_1010_LEDGER_GRP,False)   
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FB50L_1010_REFERENCE,False)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_FB50L_1010_DOCHEADER_TEXT,False) 
Call SetTextbox("Document type","ACGL_HEAD-BLART","",DT_FB50L_1010_DOCUMENT_TYPE,False) 
Call TakeScreenShot()

Call ClickButtonIfExist("Cancel   \(F12\)",True)
wait(2)

'Fill the details in Table
Call SetTableData("SAPLFSKBTABLE","G/L Acct","1","","",DT_FB50L_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableData("SAPLFSKBTABLE","D/C","1","","","Credit",False)
Call PressEnter() 
Call PressEnter() 
Call SetTableData("SAPLFSKBTABLE","Amount in doc.curr.","1","","",DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)

Call SetTableData("SAPLFSKBTABLE","G/L Acct","2","","",DT_FB50L_0100_TABLECELL_GL_ACCT_1,False)
Call SetTableData("SAPLFSKBTABLE","D/C","2","","","Debit",False)
Call PressEnter() 
Call SetTableData("SAPLFSKBTABLE","Amount in doc.curr.","2","","",DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_1,False)
Call SetTableData("SAPLFSKBTABLE","Business Area","2","","",DT_FB50L_0100_TABLECELL_BUSINESS_AREA_1,False)
Call SetTableData("SAPLFSKBTABLE","Cost Center","2","","",DT_FB50L_0100_TABLECELL_COST_CENTER_1,False)
Call TakeScreenShot()


'Click on Simulate Document Posting
Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
wait(2)
Call TakeScreenShot()

'Post the Document
Call ClickButton("Post   \(Ctrl\+S\)",False) 
Wait(1)
Call VerifyStatusBarMessageType("S")

'Validate If Document is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB50L_1001_CHECK_TEXT_OF_STATUSBAR)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

