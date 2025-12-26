

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.12.06.01.01 CO Actual Assessment
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.12.06.01.01 CO Actual Assessment"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath) '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'----------------------Tcode KSU1----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextbox("Controlling Area","SVALD-VALUE","",DT_KSU1_0300_CONTROLLING_AREA,True)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INDEX",(Cint(DT_INDEX)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call PressEnter()
Call SetTextbox("Cycle","RKAL1-KSCYC","",DT_KSU1_0102_CYCLE,False)
Call SetTextbox("Start Date","T811C-SDATE","",ConvertDate(DT_KSU1_0102_START_DATE),False)
Call TakeScreenShot()

Call FocusTextBox("Start Date","T811C-SDATE",False)
Call PressEnter()
Call SetTextbox("Text","RKAL1-CTXT","",DT_KSU1_0201_TEXT,False)
Call TakeScreenShot()
Call PressEnter()
Call ClickButton("Attach segment   \(Shift\+F8\)",fALSE)
Call SetTextbox("Segment Name","KGALS-NAME","",DT_KSU1_0300_SEGMENT_NAME,False)
Call SetTextbox("Assessment CEle","KGALS-ASACC","",DT_KSU1_0307_ASSESSMENT_CELE,False)
Call TakeScreenShot()
Call FocusTextBox("Assessment CEle","KGALS-ASACC",False)
Call SelectTab("SEG_TABSTRIP","Senders/Receivers",False)
Call TakeScreenShot()

Call SetTextbox("Cost Center","KGALK-VALMIN","",DT_KSU1_0306_COST_CENTER,False)
Call SetTextbox("Cost Element","KGALK-VALMIN","",DT_KSU1_0306_COST_ELEMENT,False)
Call SetTextbox("Cost Center","KGALK-VALMIN",1,DT_KSU1_0306_COST_CENTER_OCC1,False)

Call SelectTab("SEG_TABSTRIP","Sender Values",False)
Call TakeScreenShot()

Call SelectTab("SEG_TABSTRIP","Receiver Tracing Factor",False)
Call TakeScreenShot()

Call SetTextbox("Group","KGALK-SETID","",DT_KSU1_0461_GROUP,False)
Call FocusTextBox("Group","KGALK-SETID",False)
Call ClickButton("Header data   \(F5\)",fALSE)
Call ClickButton("No check   \(Ctrl\+S\)",fALSE)

Call GetTextStatusBar("DT_KSU1_0201_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_KSU1_0201_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_KSU1_0201_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_KSU1_0201_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Exit   \(Shift\+F3\)",fALSE)
'''----------------------Tcode KSU3---------------------------- 

Call SetTcode(DT_KSU1_0100_OKCD)
Call PressEnter() 
Call PressEnter()
Call ClickButton("First segment   \(F6\)",fALSE)
Call SelectTab("SEG_TABSTRIP","Senders/Receivers",False)
Call TakeScreenShot()

Call VerifyTextBoxContent("Cost Center", "KGALK-VALMIN","",DT_KSU1_0306_CHECK_TEXT_OF_COST_CENTER,False)
Call VerifyTextBoxContent("Cost Element", "KGALK-VALMIN","",DT_KSU1_0306_CHECK_TEXT_OF_COST_ELEMENT,False)
Call VerifyTextBoxContent("Cost Center", "KGALK-VALMIN",1,DT_KSU1_0306_CHECK_TEXT_OF_COST_CENTER_OCC1,False)
Call ClickButton("Exit   \(Shift\+F3\)",fALSE)
''''----------------------Tcode KSU5----------------------------
Call SetTcode(DT_KSU1_0100_OKCD_OCC1)
Call PressEnter() 
Call SetTextbox("Fiscal Year","RKGA2U-GJAHR","",DT_KSU1_0101_FISCAL_YEAR,False)
Call SetTextbox("Period","RKGA2U-FROM","",DT_KSU1_0101_PERIOD,False)
Call SetTextbox("To","RKGA2U-TO","",DT_KSU1_0101_TO,False)
Call TakeScreenShot()

Call SetTextbox("Cycle","RKGA2-KSCYC","",DT_KSU1_0101_CYCLE,False)
Call SetTextbox("Start Date","RKGA2-SDATE","",ConvertDate(DT_KSU1_0101_START_DATE),False)
Call ClickButton("Execute   \(F8\)",fALSE)


Call SetFocusguiLabel("Warning","","",False)
Call SendKey("{F2}")
Wait 5
Call VerifyifGuiLabelExists(DT_KSU1_0120_CHECK_TEXT_OF_NO_VALID_SENDER_ENTRIES_WERE_FOUND)
Call TakeScreenSHot()
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Exit   \(Shift\+F3\)",fALSE)
Call ClickButton("Yes",True)
Call ClickButton("Exit   \(Shift\+F3\)",fALSE)
'''----------------------Tcode S_ALR_87013611----------------------------
'
Call SetTcode(DT_KSU1_0100_OKCD_OCC2)
Call PressEnter() 
Call SetTextbox("Or value\(s\)","_1KOSET-LOW","",DT_KSU1_1000_OR_VALUE_S,False)
Call SetTextbox("Or value\(s\)","_1KSTAR-LOW","",DT_KSU1_1000_OR_VALUE_S_OCC1,False)
Call SetTextbox("From Period","\$1PERIV","",DT_KSU1_1000_FROM_PERIOD,False)
Call SetTextbox("To Period","\$1PERIB","",DT_KSU1_1000_TO_PERIOD,False)
Call TakeScreenShot()

Call FocusTextBox("Or value\(s\)","_1KSTAR-LOW",False)
Call ClickButton("Execute   \(F8\)",fALSE)

Call TakeScreenShot()

Call ClickButtonifExist("Page right   \(Ctrl\+F11\)",fALSE)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",fALSE)
Call ClickButtonifExist("Yes",True)
Call SetTextbox("Cost Center Group","\$1KOSET","",DT_KSU1_1000_COST_CENTER_GROUP,False)
Call SetTextbox("Or value\(s\)","_1KOSET-LOW","","",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",fALSE)
wait(10)
Call TakeScreenShot()
'The below validation is no lognger exists.
'Call VerifyifGuiLabelExists(DT_KSU1_0120_CHECK_TEXT_OF_G02OS1)
Call ClickButtonifExist("Page right   \(Ctrl\+F11\)",fALSE)
Call TakeScreenShot()

Call LogOff'
Call FinalStatus()''
