
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Post GL Reccuring Documents_p3
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
'.................Test Script Name : Test_Post GL Reccuring Documents_p3
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Post GL Reccuring Documents_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Post GL Reccuring Documents_p3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System

'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode F.15----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter the details
Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_F15_1000_COMPANY_CODE,False)  
Call SetTextbox("Document Number","BR_BELNR-LOW","",DT_F15_1105_DOCUMENT_NUMBER,False)  
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_F15_1000_FISCAL_YEAR,False)   
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_F15_1000_POSTING_DATE),False)   
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_F15_1000_TO),False)   

Call SelectCheckbox("KREDITOR",0,DT_F15_1000_VENDOR_ACCOUNTS,False)
Call SelectCheckbox("DEBITOR",0,DT_F15_1000_CUSTOMER_ACCOUNTS,False)
Call SelectCheckbox("NABEL",0,DT_F15_1000_NONFILLED_DOCUMENTS,False)
Call SelectCheckbox("SACHKTO",0,DT_F15_1000_DISPLAY_GL_ACCOUNTS,False)
Call SelectCheckbox("BELNL",0,DT_F15_1000_DOCUMENTS_FROM_NEXT_RUN,False)
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()


Call VerifyifGuiLabelExists(DT_F15_1105_DOCUMENT_NUMBER)
Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_F15_1000_POSTING_DATE),"wnd\[0\]/usr/lbl\[44,6\]")
''Call VerifyifGuiLabelExistsByRelativeid(Replace(DT_F15_1000_POSTING_DATE,"/","."),"wnd\[0\]/usr/lbl\[44,6\]")
''Call VerifyifGuiLabelExists(Replace(DT_F15_1000_POSTING_DATE,"/","."))
''Call VerifyifGuiLabelExists(ConvertDate(DT_F15_1000_POSTING_DATE))
Call VerifyifGuiLabelExists(Replace(DT_F15_1000_TO,"/","."))

'Click on Exit
Call ClickButton("Exit   \(Shift\+F3\)",False) 
Wait(2)
'Click on Exit
Call ClickButton("Exit   \(Shift\+F3\)",False) 
Wait(2)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

