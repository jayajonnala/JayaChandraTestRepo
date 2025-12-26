

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_SBWP_Approve_WF
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_SBWP_Approve_WF"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_002_Create_Purchasing_Conditions_for_an_Existing_Article_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =4
Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'----------------------Tcode SBWP----------------------------

'Enter the Transaction Code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Activate the Node
Call ActivateNodeGuiTree(0,"#1;#1;#3")
Wait(2)
Call TakeScreenShot()

'Click on Filtr button
Call ClickButtonToolBar("&MB_FILTER",0)
Wait 2
Call ClickButtonToolBar("&FIND", 0)


'Search the Filter item
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Work Item Title",True) 
Call TakeScreenShot()

'Click on OK
Call ClickButton("OK   \(Enter\)",True) 
Wait(2)

'Verify If Hit Is found
Call VerifyTextBoxContent("Information Message","GS_SEARCH-SEARCH_INFO",0,Lcase("Hit displayed : 1"),True)


'Click on OK
Call ClickButton("Cancel   \(F12\)",True) 
Wait(2)


Call ClickButton("Add Filter Criterion \(F7\)",True) 
Wait(2)

Call ClickButton("600_BUTTON",True) 
Wait(2)

Call SetTextbox("Work Item Title","%%DYN001-LOW","",DT_ARTICLE_DESCRIPTION,True) 
Call TakeScreenShot()

Call ClickButton("Execute   \(Enter\)",True) 

'Approval/Enrichment for Article: AUTO_EN_TEST216 required 
Call SelectRowGuiGrid("Workflow.*",0,"Title",DT_SBWP_1000_GRIDCELL_1_WORK_ITEM_TITLE,False)
Call DoubleClick()
Wait(1)
Call TakeScreenShot()
Call ClickButton("Submit completed Form to PRICAT",False) 
Wait(2)

Call TakeScreenShot()

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

