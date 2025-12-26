
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrTestCaseName = "Test_POST_03-01-01-01-Cre Venr"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-------------------------XK02--------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()  
Call TakeScreenShot

Call SelectCheckbox("RF02K-D0120","5","ON",false)
Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK02_0101_VENDOR,False)   
'This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK02_0101_COMPANY_CODE,False)
Call SetTextbox("Purchasing Organization","RF02K-EKORG","",DT_XK02_0101_PURCH_ORGANIZATION,False)    
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
call SetTextbox("VAT Reg\. No\.","LFA1-STCEG","","",false)  '''No need to give VAT No
call PressEnter
Call TakeScreenShot
call ClickButtonIfExist("btn\[11\]",false)	''save button
Call TakeScreenShot
call PressEnter
Call VerifyStatusBarMessageType("S")
call SelectCheckbox("RF02K-D0210","5","ON",false)
Call TakeScreenShot
call PressEnter
call PressEnter
call PressEnter
Call TakeScreenShot
call SetTextbox("Prev\.acct no\.","LFB1-ALTKN","",DT_EUP_NO,false)  
call ClickButtonIfExist("btn\[11\]",false)
Call TakeScreenShot
Call VerifyStatusBar("Changes have been made")
' SelectCheckbox(checkboxName, checkBoxIndex, OnOffStatus, blnIsItPopup)

Call LogOff()
Call FinalStatus ()
