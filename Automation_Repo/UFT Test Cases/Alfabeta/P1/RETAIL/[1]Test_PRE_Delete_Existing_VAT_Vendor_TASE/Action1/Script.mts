

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_Delete_Existing_VAT_Vendor
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


gstrTestCaseName = "Test_PRE_Delete_Existing_VAT_Vendor"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.04.02.21 VIM - PO Precontrole Issue - BR01 - Invalid Company_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------------------  XK02------------------------------------------------
Call SetTextboxNoLabel("RF02K-LIFNR","",DT_VENDOR_NO,False)
Call SelectCheckbox("RF02K-D0120",0,DT_XK02_0101_CONTROL,False)
Call TakeScreenShot()
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Tax Number 2","LFA1-STCD2",0,FormatBlank(DT_XK02_0120_TAX_NUMBER_2),False)
Call SetTextbox("Tax Number 1","LFA1-STCD1",0,"",False)
Call SetTextbox("VAT Reg\. No\.","LFA1-STCEG",0,FormatBlank(DT_XK02_0120_VAT_REG_NO),False)
Call SetTextbox("Location no\. 1","LFA1-BBBNR",0,"",False)
Call SetTextbox("Location no\. 2","LFA1-BBSNR",0,"",False)
Call SetTextbox("Check digit","LFA1-BUBKZ",0,"",False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call PressEnter()     ' 
Call TakeScreenShot()
'Call VerifyStatusBarMessageType("S")
'Call VerifyStatusBar(DT_XK02_0101_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()



