
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.4.1.2.5. Listing at DC
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
'.................Test Script Name :Test_2.4.1.2.5. Listing at DC
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.4.1.2.5. Listing at DC"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.4.1.2.5. Listing at DC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System

'DataRowSet=3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
'SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'----------------------Tcode WAK1----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_MM42_0100_PURCHASING_ORG,False)
'Call SetTextbox("Vendor","RMMW1-LIFNR","",DT_MM42_0100_VENDOR,False)
Call SetTextboxNoLabel("RMMW1-LIFNR","",DT_MM42_0100_VENDOR,False)
Call TakeScreenShot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", "Basic Data", False)

Call PressEnter()

Call ClickButton("Go to additional data   \(Ctrl\+F6\)",False)
Call TakeScreenShot()

Call SelectTab("TABSPR1","Layout Modules", false)
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLWLAY_DIA_ARTTC_LAY","Alt. Unit",4,DT_MM42_8024_TABLECELL_ALT_UNIT_1, False)
Call SetTableDataNoRef("SAPLWLAY_DIA_ARTTC_LAY","Layout Module",4,DT_MM42_8024_TABLECELL_LAYOUT_MODULE_1, False)
Call SetTableDataNoRef("SAPLWLAY_DIA_ARTTC_LAY","Facing",4,DT_MM42_8024_TABLECELL_FACING_1, False)

Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",False)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()

Call SetTcode(DT_MM42_0100_OKCD)
Call PressEnter()


Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)", False)
Call TakeScreenShot()
Call SetTextbox("Created By","ENAME-LOW","","",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)", True)

'''' ActivateCellGuiGridByRefVal(gridTitle, gridIndex, refColumn, refFieldVal, columnName, blnIsItPopup)
'''Call ActivateCellGuiGridByRefVal

Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RWSORT07L",0,10,True)
Call TakeScreenShot()
Call ClickButton("Choose   \(F2\)",True)
wait 1
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot()


Call ActivateNodeGuiTree(0,"#2;#1")
Call TakeScreenShot()


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

