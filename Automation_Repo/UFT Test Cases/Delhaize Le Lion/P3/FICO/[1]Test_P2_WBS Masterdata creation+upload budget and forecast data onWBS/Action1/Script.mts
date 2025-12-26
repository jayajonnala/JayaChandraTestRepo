
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
gstrTestCaseName = "Test_P2_WBS Masterdata creation+upload"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Login 
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-------------------------VL10G----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Project def\.","\*PROJ-PSPID","",DT_CJ02_0100_PROJECT_DEF,False)
Call TakeScreenShot
Call ClickButton("WBS element overview   \(F5\)",False)
Call SetTableData("SAPLCJWBTAB_902","Level",2,"","",DT_CJ02_0902_TABLECELL_LEVEL_1,False)
Call SetTableData("SAPLCJWBTAB_902","Level",3,"","",DT_CJ02_0902_TABLECELL_LEVEL_2,False)
Call TakeScreenShot
Call SetTableData("SAPLCJWBTAB_902","WBS element",1,"","",DT_CJ02_0902_TABLECELL_WBS_ELEMENT_0,False)
Call SetTableData("SAPLCJWBTAB_902","WBS element",2,"","",DT_CJ02_0902_TABLECELL_WBS_ELEMENT_1,False)
Call SetTableData("SAPLCJWBTAB_902","WBS element",3,"","",DT_CJ02_0902_TABLECELL_WBS_ELEMENT_2,False)
Call TakeScreenShot
Call SetTableData("SAPLCJWBTAB_902","Description",1,"","",DT_CJ02_0902_TABLECELL_DESCRIPTION_0,False)
Call SetTableData("SAPLCJWBTAB_902","Description",2,"","",DT_CJ02_0902_TABLECELL_DESCRIPTION_1,False)
Call SetTableData("SAPLCJWBTAB_902","Description",3,"","",DT_CJ02_0902_TABLECELL_DESCRIPTION_2,False)
Call TakeScreenShot
Call SetTableData("SAPLCJWBTAB_902","Typ",1,"","",DT_CJ02_0902_TABLECELL_TYP_0,False)
Call SetTableData("SAPLCJWBTAB_902","Typ",2,"","",DT_CJ02_0902_TABLECELL_TYP_1,False)
Call SetTableData("SAPLCJWBTAB_902","Typ",3,"","",DT_CJ02_0902_TABLECELL_TYP_2,False)
Call TakeScreenShot
Call SelectTab("TABCJLE","Responsibilities",False)
Call TakeScreenShot
Call SetTableData("SAPLCJWBTAB_904","Req.cost center",1,"","",DT_CJ02_0904_TABLECELL_REQCOST_CENTER_0,False)
Call SetTableData("SAPLCJWBTAB_904","Req.cost center",2,"","",DT_CJ02_0904_TABLECELL_REQCOST_CENTER_1,False)
Call TakeScreenShot
Call SelectRowGuiTable("SAPLCJWBTAB_904","Lev","3",False)
Call SelectMenuBar("Extras;Investment program")
Call TakeScreenShot
Call SetTextbox("Investment program","RAIP1-PRNAM","",DT_CJ02_0050_INVESTMENT_PROGRAM,true)
Call SetTextbox("Position ID","RAIP1-POSID","",DT_CJ02_0050_POSITION_ID,true)
Call SetTextbox("Approval year","RAIP1-GJAHR","",DT_CJ02_0050_APPROVAL_YEAR,true)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot
Call Clickbutton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_CJ02_0100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

Call LogOff()
Call FinalStatus()

