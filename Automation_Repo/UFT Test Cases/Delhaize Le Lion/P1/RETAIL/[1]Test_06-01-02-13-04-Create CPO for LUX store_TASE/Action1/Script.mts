
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.46 VIM - PO Precontrole Issue - BR24 - Missing Posting Date (PO)
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


gstrTestCaseName = "Test_06-01-02-13-04-Create CPO for LUX store"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'all LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()
''
'--------------------------------------------  ZMDPU_COLLECT_PO----------------------------------------------


Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)

Call SetTextbox("Variant","V-LOW","",DT_ZMDPU_COLLECT_PO_0100_VARIANT,False)
Call SetTextbox("Created By","ENAME-LOW","","",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",true)
Call TakeScreenShot()
'
'''line 56 added by KGARA on 03/06/2022'
Call SetTextbox("Sales Order","S_VBELN-LOW","",DT_SALES_ORD,False)
'Call SetTextbox("Sales Order","S_VBELN-LOW","","",False)
Call SetTextbox("Stock Transfer Order Document","S_EBELN-LOW","",DT_ZMDPU_COLLECT_PO_1000_STOCK_TRANSFER_ORDER_DOCUMENT,False)
Call ClickButton("Execute   \(F8\)",false)
Call TakeScreenShot()
Call GetTextStatusBar("DT_ZMDPU_COLLECT_PO_0500_CHECK_TEXT_OF_STATUSBAR")

Call TakeScreenShot()
'''No need to click create CPO button as per TAO log
Call ClickButtonIfExist("Create CPO   \(F7\)",false)
Call GetGridContent("", 0, "Purchasing Document", 1, "<NA>", "<NA>", "DT_ZMDPU_COLLECT_PO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CPO_NUMBER_OUTPUT")
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()


