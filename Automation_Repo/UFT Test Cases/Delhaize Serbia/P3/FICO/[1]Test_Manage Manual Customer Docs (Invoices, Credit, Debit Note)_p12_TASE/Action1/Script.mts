
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Manual Customer Docs (Invoices, Credit, Debit Note)_p12
'.................Author : TCS        :Jaya
'................ Creation Date    : 
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Manual Customer Docs (Invoices, Credit, Debit Note)_p12_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear GL Accounts  Manual and Automatic_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''		


'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-FBL5N ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("X_NORM", 0, DT_FBL5N_1000_NORMAL_ITEMS, False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)
Call SetTextbox("Open at key date","PA_STID2","",ConvertDate(DT_FBL5N_1000_OPEN_AT_KEY_DATE),False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickLabel("DocumentNo", 0, False)
Call TakeScreenShot
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButton("Multiple selection",True)

Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC2, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "", "", DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC3, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 5, "", "", DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC4, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 6, "", "", DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC5, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 7, "", "", DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC6, True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call ClickLabel("DocumentNo", 0, False)
Call TakeScreenShot
Call ClickButton("Sort in ascending order   \(Ctrl\+Shift\+F5\)",False)
Call TakeScreenShot

Call VerifyifGuiLabelExists(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0)
Call VerifyifGuiLabelExistsByRelativeid(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART,"wnd\[0\]/usr/lbl\[82,10\]")
Call VerifyifGuiLabelExists(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC1)
Call VerifyifGuiLabelExists(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyifGuiLabelExists(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC2)
Call VerifyifGuiLabelExists(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
Call VerifyifGuiLabelExists(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC3)
Call VerifyifGuiLabelExistsByRelativeid(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLART,"wnd\[0\]/usr/lbl\[82,13\]")
Call VerifyifGuiLabelExists(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC4)
Call VerifyifGuiLabelExistsByRelativeid(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BLART,"wnd\[0\]/usr/lbl\[82,14\]")
Call VerifyifGuiLabelExists(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC5)
Call VerifyifGuiLabelExistsByRelativeid(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BLART,"wnd\[0\]/usr/lbl\[82,15\]")
Call VerifyifGuiLabelExists(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC6)
Call VerifyifGuiLabelExistsByRelativeid(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BLART,"wnd\[0\]/usr/lbl\[82,16\]")

'''Clicking label will open the document overview
Call ClickLabel(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0, 0, False)
TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
TakeScreenShot
Call ClickButton("Back   \(F3\)",False)

Call ClickLabel(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC1, 0, False)
TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
TakeScreenShot
Call ClickButton("Back   \(F3\)",False)

Call ClickLabel(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC2, 0, False)
TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
TakeScreenShot
Call ClickButton("Back   \(F3\)",False)

Call ClickLabel(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC3, 0, False)
TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
TakeScreenShot
Call ClickButton("Back   \(F3\)",False)

Call ClickLabel(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC4, 0, False)
TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
TakeScreenShot
Call ClickButton("Back   \(F3\)",False)

Call ClickLabel(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC5, 0, False)
TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
TakeScreenShot
Call ClickButton("Back   \(F3\)",False)

Call ClickLabel(DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0_OCC6, 0, False)
TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
TakeScreenShot

Call ClickButton("Exit   \(Shift\+F3\)",False)
TakeScreenShot
Call ClickButton("Exit   \(Shift\+F3\)",False)
TakeScreenShot

Call LogOff()
Call FinalStatus()
