'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Selection variant maintenance - Internal Order
'.................Author : TCS        :Jaya
'................ Creation Date    : 
'.................Modified By :
'.................Modified Date/Details :

'''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


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
'''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_Selection variant maintenance - Internal Order"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear GL Accounts  Manual and Automatic_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'

'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''----------------------Login-------------------------'''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
Call TakeScreenShot()
'''----------------------Tcode OKOV----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Variant","RSVAR-VARIANT","",(DT_OKOV_0304_CREATE_VARIANT),False)
Call ClickBUtton("Create",False)
Call TakeScreenShot()

Call SetTextbox("Order type","AUART-LOW","",(DT_OKOV_1000_ORDER_TYPE),False)
Call SetTextbox("Company code","BUKRS-LOW","",(DT_OKOV_1000_COMPANY_CODE),False)
Call TakeScreenShot()
Call ClickBUtton("Attributes   \(F6\)",False)
Call TakeScreenShot()
Call SetTextbox("Description","RSVAR-VTEXT","",(DT_OKOV_0281_DESCRIPTION),False)

Call ClickBUtton("Save   \(Ctrl\+S\)",False)

Call GetTextStatusBar("DT_OKOV_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_OKOV_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_OKOV_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_OKOV_1000_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus()

