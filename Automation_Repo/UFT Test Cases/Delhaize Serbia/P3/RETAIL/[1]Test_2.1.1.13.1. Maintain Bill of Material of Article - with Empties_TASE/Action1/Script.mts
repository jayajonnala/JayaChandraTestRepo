
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_2.1.1.13.1. Maintain Bill of Material of Article - with Empties
'.................Author : TCS 	   :
'................ Creation Date    :
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

gstrTestCaseName = "Test_2.1.1.13.1. Maintain Bill of Material of Article - with Empties"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.1.1.1. Create Article Purchasing Data.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=7
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''----------------------Tcode-MM42 ----------------------------''''''
'Enter the transaction code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_MM42_0100_PURCHASING_ORG,False)
Call SetTextboxNoLabel("RMMW1-LIFNR","",DT_MM42_0100_VENDOR,False)

Call TakeScreenShot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Purchasing",False)
Call TakeScreenShot()
Call PressEnter()
wait(2)
Call TakeScreenShot()

Call ClickButton("Goto Component Creation",False)
Call TakeScreenShot()
Call SetTextbox("Unit of Measure","WSTR_DYNP-MEINH","",DT_MM42_0100_UNIT_OF_MEASURE,False)
Call PressEnter()
wait(2)
Call TakeScreenShot()

Call SetTableData("SAPLWST1TC_COMPONENTS", "Component", 1, "", "", DT_MM42_0100_TABLECELL_COMPONENT_0, False)
Call SetTableData("SAPLWST1TC_COMPONENTS", "Component qty", 1, "", "", DT_MM42_0100_TABLECELL_COMPONENT_QTY_0, False)
Call PressEnter()
wait(2)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Wait(2)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Goto Component Creation",False)
Call TakeScreenShot()
Call SetTextbox("Unit of Measure","WSTR_DYNP-MEINH","",DT_MM42_0100_UNIT_OF_MEASURE_OCC1,False)
Call PressEnter()
Call TakeScreenShot()

Call VerifyTableCellContent(1,"Component" , "SAPLWST1TC_COMPONENTS", DT_MM42_0100_CHECK_TEXT_OF_TABLECELL_COMPONENT_0)

Call LogOff()
Call FinalStatus ()

