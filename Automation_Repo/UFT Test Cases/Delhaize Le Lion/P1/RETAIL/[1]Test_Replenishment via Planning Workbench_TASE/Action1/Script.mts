
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : [1]Test_Replenishment via Planning Workbench
'.................Test Scenario: AT_P2P_Fresh_Replenishment via Workbench and Auction PO
'.................TCode:WWP1; ME23N
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

gstrTestCaseName = "Test_Replenishment via Planning Workbench"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
'
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode WWP1----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Click on Get Variant
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Wait(2)
'clear Created By
Call SetTextbox("Created By","ENAME-LOW",""," ",FALSE)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
'click on row with BW09_006
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RWWP001"&".*",1,12,False)
Call TakeScreenShot()

Call ClickButton("Choose   \(F2\)",False)
Call PressEnter()
'click on New entry
Call ClickButton("Execute   \(F8\)",False)
Call PressEnter()

Call SelectNodeGuiTree(0, "#1;#5;#1;#1;#1")
Call ClickButtonToolbar("EXE",0)

For i = 1 To 1000
Call GetTextboxValue("G_KOPF_0100_EBELN_ANZEIGE", 0, "DT_PONUMBER_VAL", False)
If DT_PONUMBER_VAL <> "New entry" Then
	Call ClickButton("Back   \(F3\)",False)
	Wait 2
	Call ClickBUttonIfExist("Yes",True)
	Call SelectNodeGuiTree(0, "#1;#5;#1;#1;#"&i&"")
	Call ClickButtonToolbar("EXE",0)
Else 
	Exit for	
End if
Next

Call SetTableData("SAPLWOD1CONTROL_ARTICLES_TC","PO Quantity",1,"OUn","CV",DT_PO_QUANTITY_0,False)

Call TakeScreenShot()
Call ClickButtonIfExist("Release purchase order   \(Shift\+F11\)",False)
wait(30)

Call SetTextbox("Release Code","G_FRGCO_0500","","01",True)
Call PressEnter()

Call GetTextboxValue("G_KOPF_0100_EBELN_ANZEIGE",0,"DT_PO_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_PO_OUTPUT",DT_PO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()


'------------------ME23N-------------------------------'

Call SetTcode(DT_WWP1_0100_OKCD) 
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_PO,True)
Call PressEnter()
Call TakeScreenShot()

' VerifyComboBoxValue(comboAttachedText, expectedValue)
'Call VerifyComboBoxValue("MEPO_TOPLINE-BSART",DT_WWP1_1105_CHECK_VALUE_OF_MEPO_TOPLINEBSART)



'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


