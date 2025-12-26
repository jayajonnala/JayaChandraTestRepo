
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.1.1.12.1. Maintain,Check Purchasing Conditions of Article - Create
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
'.................Test Script Name : Test_2.1.1.12.1. Maintain,Check Purchasing Conditions of Article - Create
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.1.1.12.1. Maintain,Check Purchasing Conditions of Article - Create"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.1.1.12.1. Maintain,Check Purchasing Conditions of Article - Create.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode MM42 ----------------------------
'Enter the transaction code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the details
Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call SetTextboxNoLabel("RMMW1-LIFNR","",DT_MM42_0100_VENDOR,False)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_MM42_0100_PURCHASING_ORG,False)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG,False)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL,False)

'Capture the screenshot
Call TakeScreenShot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Purchasing",False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
wait(2)

'Navigate to Purchasing Tab
Call SelectTab("TABSPR1"," Purchasing",False)
Wait(1)
Call TakeScreenShot()

'Click on Conditions Button
Call ClickButton("For the Conditions",False)
Wait(2)
Call TakeScreenShot()

'Get the valid From Date
Call GetTableCellData("SAPLV14ATCTRL_D0102","Valid From",1,"Valid to","31.12.9999","DT_MM42_0201_VALID_FROM",False)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Click on New validity period with reference  Button
Call ClickButton("New validity period with reference   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

'Set New valid From Date
Call SetTextbox("Valid From","RV13A-DATAB","",ConvertDate(DT_MM42_0201_VALID_FROM),False)
Call TakeScreenShot()
Call PressEnter()
Wait(2)

'enter the details
Call SetTableData("SAPMV13ATCTRL_D0201","CnTy",2,"","",DT_MM42_0201_TABLECELL_CNTY_2,False)
Call SetTableData("SAPMV13ATCTRL_D0201","Amount",2,"","",DT_MM42_0201_TABLECELL_AMOUNT_2,False)
Call PressEnter()
Wait(2)
Call TakeScreenShot()


'Save the Details
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()


'Click on Enterif pop up appears
Call ClickButtonIfExist("Enter   \(F5\)",True)
Wait(2)


'Verify the Message Type
Call VerifyStatusBarMessageType("S")

'Verify the status bar message
Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

