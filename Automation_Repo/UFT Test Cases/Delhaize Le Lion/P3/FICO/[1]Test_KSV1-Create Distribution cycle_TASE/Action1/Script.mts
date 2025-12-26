
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'----------------------------------------------------------------------------------------------------------------------------
gstrTestCaseName = "Test_KSV1-Create Distribution cycle_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Login 
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-------------------------VL10G----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Controlling Area","SVALD-VALUE","",DT_KSV1_0300_CONTROLLING_AREA,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Cycle","RKAL1-KSCYC","",DT_KSV1_0102_CYCLE,False)
Call TakeScreenShot
Call SetTextbox("Start Date","T811C-SDATE","",ConvertDate(DT_KSV1_0102_START_DATE),False)
Call SetTextbox("Cycle","RKAL1-RKSCYC","",DT_KSV1_0102_CYCLE_OCC1,False)
Call TakeScreenShot
Call SetTextbox("Start Date","RKAL1-RSDATE","",DT_KSV1_0102_START_DATE_OCC1,False)
Call SetTextbox("Controlling Area","RKAL1-RKOKRS","",DT_KSV1_0102_CONTROLLING_AREA,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxContent("Status","RKAL1-STATUS","",DT_KSV1_0201_CHECK_TEXT_OF_STATUS,False)
Call SetTextbox("To","T811C-EDATE","",ConvertDate(DT_END_DATE),False)
Call TakeScreenShot
Call SetTextbox("Text","RKAL1-CTXT","",DT_KSV1_0201_TEXT,False)
Call TakeScreenShot
Call ClickButton("No check   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_KSV1_0201_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus()


'Call SetTextbox("Project def\.","PROJ-PSPID","",DT_CJ30_0200_PROJECT_DEF,False)
'Call TakeScreenShot
'Call ClickButton("Original Budget   \(F5\)",False)
'Call SetTableData("SAPLKBPPTC_320","Budget",1,"","",DT_CJ30_0320_TABLECELL_BUDGET_0,False)
'Call SetTableData("SAPLKBPPTC_320","Budget",2,"","",DT_CJ30_0320_TABLECELL_BUDGET_1,False)
'Call SetTableData("SAPLKBPPTC_320","Budget",3,"","",DT_CJ30_0320_TABLECELL_BUDGET_2,False)
'Call TakeScreenShot
'Call ClickButton("Save   \(Ctrl\+S\)",false)
'Call GetStatusBar("item1","DT_CJ30_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
'Call WriteRunTimeDataToExcelGlobalSheet("DT_CJ30_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_CJ30_0200_CHECK_TEXT_OF_STATUSBAR)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call VerifyStatusBar("Document "&DT_CJ30_0200_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT&" posted")
'
''----------------/nS-ALR-180358777--------------------------'
'
'Call SetTcode(DT_CJ30_0200_OKCD)     
'Call PressEnter()     
'Call TakeScreenShot
'Call SendKey("{F4}")
'Call SendKey("{F2}")
'Call Settextbox("Project","CN_PROJN-LOW","",DT_CJ30_1000_PROJECT,False)
'Call TakeScreenShot
'Call ClickButton("Execute   \(F8\)",False)
'Call TakeScreenShot
'Call VerifyNodeTextGuiTree(0,"#2")
'



