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


'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p15_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 1st May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p15_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Manual Vendor Invoicing  Crediting (Non PO Invoices)_p15_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FBL1N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company code","KD_BUKRS-LOW","",DT_FBL1N_1000_COMPANY_CODE,False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FBL1N_1000_VENDOR_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Company code","KD_BUKRS-LOW",False)
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
'
'''----------------------Tcode FBL3N----------------------------
'Enter the Tcode
Call SetTcode(DT_FBL1N_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FBL1N_1000_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FBL1N_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FBL1N_1000_GL_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Company code","SD_BUKRS-LOW",False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
Call ClickButton("Execute   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
CAll ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(10)
Call VerifyStatusBarMessageType("S")
'Capture the screenshot
Call TakeScreenShot()

Call ClickLabel("DocumentNo","", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL3N_1105_DOCUMENT_NUMBER,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
Wait(3)
'Capture the screenshot
Call TakeScreenShot()


Function VerifyifGuiLabelWithIndexExists(Content,Index)
	
 If Not (Environment.Value("blnFatalError") or Content= DS_SKIP) Then
	If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : VerifyifGuiLabelWithIndexExists"
	
	strStepName = "Verify if Gui Label exists with index"

    If Content <>"" Then
     set   objLabel = SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiLabel("content:="&Content,"guicomponenttype:=30","index:="&Index)
               If objLabel.Exist Then
					Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","2",Content,"Gui Label with value "&Content &"index : "&Index &" exists in the screen")	
					strStatus = "PASS"
					strMsg = "Gui Label with value "&Content&"index : "&Index &" exists in the screen"	
					blnCaptureFlag = True
			              	If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
		            	  		ImagePath=CaptureScreenshot(strStepName,objLabel,False,False,False)
                              End If
					
					Else
					Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","1"," Status bar Content","Gui Label with value "&Content &"index : "&Index &" doesn't exist  in the screen")	
					strStatus = "FAIL"
					blndefectFlag =True
					strMsg = "Gui Label with value "&Content&" doesn't exist  in the screen"
					blnObjectError=True
				End If
       	Else
    	Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","1","Gui Label","Function Parameter Not Passed Properly. Check the --VerifyifGuiLabelWithIndexExists-- Function Call")
			strStatus = "FAIL" 
			strMsg = "Function Parameter Not Passed Properly. Check the --VerifyifGuiLabelWithIndexExists-- Function Call-"
	End if


If strStatus = "FAIL"  Then
		VerifyifGuiLabelWithIndexExists = strMsg
		blnMainFailFlag = True
		ImagePath=CaptureScreenshot(strStepName,objLabel,False,False,False)
    
	Else
		VerifyifGuiLabelWithIndexExists = True
	End If
	If blnDefault_eSwiftReporting Then  
		Call UpdateResultHtml(strStepName,Content,strMsg,strStatus,"")
	End If

End If
End Function

Call VerifyifGuiLabelExists(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_56_ZUONR)
Call VerifyifGuiLabelExists(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_56_BELNR)
CAll VerifyifGuiLabelWithIndexExists(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_56_ICO_AUGP,0)

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

