'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage AP Bill of Exchange_p5_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 12th March
'.................Modified By :
'.................Modified Date/Details :
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrTestCaseName = "Test_Manage AP Bill of Exchange_p5_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AP Bill of Exchange_p5_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FEBAN----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("House bank","SL_HBKID-LOW","",DT_FEBAN_0060_HOUSE_BANK,True)
Call SetTextbox("Company Code","SL_BUKRS-LOW","",DT_FEBAN_0060_COMPANY_CODE,True)
Call SetTextbox("Account ID","SL_HKTID-LOW","",DT_FEBAN_0060_ACCOUNT_ID,True)
Call SetTextbox("Statement Date","SL_AZDAT-LOW","",Replace(DT_FEBAN_0060_STATEMENT_DATE,"/","."),True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Tree or List Display   \(Shift\+F2\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'Click on "change layout" option in "Layout / varriant" menu in grid toolbar
Call SelectMenuIdToolBar("&COL0",0)
'Capture the screenshot
Call TakeScreenShot()

'adding "Customer" and "Statement Nummber"
Call SelectCellGuiGrid("Column Set","",DT_FEBAN_0620_GRIDCELL_11_COLUMN_NAME,"Column Name",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call SelectCellGuiGrid("Column Set","",DT_FEBAN_0620_GRIDCELL_15_COLUMN_NAME_STATEMENT_NUMBER,"Column Name",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Transfer   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

'SetCurrentCell ------------------
Call SelectColumnGuiGrid("Statement Items","","Customer",False)
Call ClickButtonToolBar("&SORT_ASC",0)
'Capture the screenshot
Call TakeScreenShot()

'getting session details for future tree-node navigation
Call GetGridContentByTitle("Statement Item","","Account Currency",1,"ACCOUNT_CURRENCY_OUTPUT")
Call GetGridContentByTitle("Statement Item","","Amount",1,"AMOUNT_OUTPUT")
Call GetGridContentByTitle("Statement Item","","Statement number",1,"STATEMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'set customer nummber
Call SetGridData("Statement Items",1,"Customer",DT_FEBAN_0100_GRIDCELL_0_CUSTOMER,False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Yes",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
'''----------------------Tcode FEBAN----------------------------
'
'Enter the Tcode
Call SetTcode(DT_FEBAN_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FEBAN_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("House bank","SL_HBKID-LOW","",DT_FEBAN_0060_HOUSE_BANK_OCC1,True)
Call SetTextbox("Company Code","SL_BUKRS-LOW","",DT_FEBAN_0060_COMPANY_CODE_OCC1,True)
Call SetTextbox("Account ID","SL_HKTID-LOW","",DT_FEBAN_0060_ACCOUNT_ID_OCC1,True)
Call SetTextbox("Statement Date","SL_AZDAT-LOW","",Replace(DT_FEBAN_0060_STATEMENT_DATE_OCC1,"/","."),True)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
wait(1)
'Capture the screenshot
Call TakeScreenShot()
'Validate If doc is generated
Call GetStatusBar("item1","DT_FEBAN_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FEBAN_0100_CHECK_TEXT_OF_STATUSBAR)

'Creating node path from datasheet
Dim node_path_x
node_path_x = DT_FEBAN_0060_HOUSE_BANK_OCC1&";"&DT_FEBAN_0060_ACCOUNT_ID_OCC1&";"&DT_STATEMENT_NUMBER&"     "&DT_AMOUNT&";"&DT_FEBAN_0060_ACCOUNT_ID_OCC1&"  "&DT_ACCOUNT_CURRENCY&" "&DT_AMOUNT
'selecting appropiate node
Call ActivateNodeGuiTree("",node_path_x)
'''''Call ActivateNodeGuiTree("",DT_NODE)

'Capture the screenshot
Call TakeScreenShot()

'Double click on Doc number
SAPGuiSession("transaction:="&DT_SAPTRANSACTIONCODE).SAPGuiWindow("transaction:="&DT_SAPTRANSACTIONCODE).SAPGuiEdit("attachedtext:=Doc\.","name:=D2201_BELNR").SetFocus
Call DoubleClick()
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'validate grid components
call VerifyGridCellContent("",1,"Description","",Left(DT_FEBAN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ,4))
call VerifyGridCellContent("",1,"Posting key","",DT_FEBAN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"Description","",DT_FEBAN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)
call VerifyGridCellContent("",2,"Posting key","",DT_FEBAN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

