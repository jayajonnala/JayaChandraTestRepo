
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

gstrTestCaseName = "Test_POST_03-02-01-01-02-Create Customer in Retail"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-----------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()
Call TakeScreenShot

call SetTextboxNoLabel("RF02D-KUNNR",0,DT_XD02_7101_CUSTOMER,true)
call SetTextbox("Company code","RF02D-BUKRS","",DT_XD02_7101_COMPANY_CODE,true)
Call TakeScreenShot
call SetTextbox("Sales Organization","RF02D-VKORG","",DT_XD02_7101_SALES_ORGANIZATION,true)
call SetTextbox("Distribution Channel","RF02D-VTWEG","",DT_XD02_7101_DISTRIBUTION_CHANNEL,true)
call SetTextbox("Division","RF02D-SPART","",DT_XD02_7101_DIVISION,true)
Call TakeScreenShot()
call PressEnter
Call TakeScreenShot
Call SelectTab("TABSTRIP100","Address",False)
' GetTextboxValue(textboxName, textboxIndex, dataTableColumnName, blnIsItPopup)
Call TakeScreenShot
call GetTextboxValue("ADDR1_DATA-NAME1",1,"DT_Name_Output",false)
call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_Name_Output & "1",false)

Call TakeScreenShot
Call SelectTab("TABSTRIP100","Control Data",False)
Call TakeScreenShot
call SetTextbox("VAT Reg\. No\.","KNA1-STCEG","","",false)
Call TakeScreenShot
call ClickButton("Save   \(Ctrl\+S\)",false)
Call TakeScreenShot
Call verifystatusbar(DT_SUCCESS)
call ClickButton("btn\[0\]",false)
Call TakeScreenShot
Call LogOff()
Call FinalStatus ()
