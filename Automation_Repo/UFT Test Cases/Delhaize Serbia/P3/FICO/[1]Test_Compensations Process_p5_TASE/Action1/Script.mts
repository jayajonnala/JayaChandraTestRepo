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

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Compensations Process_p5_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 25th March
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Compensations Process_p5_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Compensations Process_p5_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'''Login'''
'DataRowSet=2
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode FB12----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Company Code","RF022-BUKRS","",DT_FB12_0100_COMPANY_CODE,False)
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call SetTextbox("Document Number","RF022-BELNR","",DT_FB12_1001_DOCUMENT_NUMBER,True)
Call SetTextbox("Fiscal Year","RF022-GJAHR","",DT_FB12_1001_FISCAL_YEAR,True)


Call ClickButton("Continue   \(Enter\)",True)
Wait(1)
Call TakeScreenShot()


Call VerifyStatusBarMessageType("S")
Call TakeScreenShot()

Call GetTextStatusBar("DT_FB12_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_FB12_0100_CHECK_TEXT_OF_STATUSBAR)

'----------------------Tcode F.64----------------------------
Call SetTcode(DT_FB12_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FB12_0100_OKCD)
Call TakeScreenShot()

Call SetTextbox("Correspondence","EVENT-LOW","",DT_FB12_1000_CORRESPONDENCE,False)
Call SetTextbox("Company code","BUKRS-LOW","",DT_FB12_1000_COMPANY_CODE,False)
Call SetTextbox("Document number","BELNR-LOW","",DT_FB12_1000_DOCUMENT_NUMBER,False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(1)
Call TakeScreenShot()


Call SelectMenuBar("Edit;Find string")
Call TakeScreenShot()

Call SetTextbox("Find","RSYSF-STRING","",DT_FB12_1000_DOCUMENT_NUMBER,True)
Call TakeScreenShot()

Call ClickButton("Find   \(Enter\)",True)
Wait(1)
Call TakeScreenShot()


Call VerifyifGuiLabelExists(DT_FB12_0120_CHECK_TEXT_OF_NO_NAME)

Call ClickLabel_NolabelContent(0, True)
'Call SetFocusGuiLabel(DT_FB12_1000_DOCUMENT_NUMBER,"88","40",True)

'Call TakeScreenShot()

'Call ClickButton("Position cursor   \(F2\)",True)

Wait(1)
Call TakeScreenShot()

Call DoubleClick()
Wait(1)
Call TakeScreenShot()

Call SetTextbox("Output Device","USR01-SPLD","",DT_FB12_1100_OUTPUT_DEVICE,True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Wait(3)
Call TakeScreenShot()

Call SetTextbox("Output Device","USR01-SPLD","",DT_FB12_1100_OUTPUT_DEVICE_OCC1,True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Wait(4)
Call TakeScreenShot()

Wait(2)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()



