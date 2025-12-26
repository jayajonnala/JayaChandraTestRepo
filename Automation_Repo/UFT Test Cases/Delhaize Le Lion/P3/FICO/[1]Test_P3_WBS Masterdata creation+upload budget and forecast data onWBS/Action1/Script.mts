
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
gstrTestCaseName = "Test_P3_WBS MD creation+upload"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''-------------------------CJ30----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Controlling Area","SVALD-VALUE","",DT_CJ30_0300_CONTROLLING_AREA,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call SetTextbox("Project def\.","PROJ-PSPID","",DT_CJ30_0200_PROJECT_DEF,False)
Call TakeScreenShot
Call ClickButton("Original Budget   \(F5\)",False)

Call SetTableData("SAPLKBPPTC_320","Budget",1,"","",DT_CJ30_0320_TABLECELL_BUDGET_0,False)
Call SetTableData("SAPLKBPPTC_320","Budget",2,"","",DT_CJ30_0320_TABLECELL_BUDGET_1,False)
Call SetTableData("SAPLKBPPTC_320","Budget",3,"","",DT_CJ30_0320_TABLECELL_BUDGET_2,False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",false)

Call GetStatusBar("item1","DT_CJ30_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_CJ30_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_CJ30_0200_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
wait 5
'''Call VerifyStatusBar("Document "&DT_CJ30_0200_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT&" posted")
Call VerifyStatusBar(DT_CJ30_0200_CHECK_TEXT_OF_STATUSBAR_OCC1)

''''-------------------------S_ALR_87013558----------------------------

Call SetTcode(DT_CJ30_0200_OKCD)     
Call PressEnter()     
Call TakeScreenShot
'''Call SendKey("{F4}")
'''Call SendKey("{F2}")

'' SetTextboxPopupIfExist(textboxName, attachedText, strTextboxValue)
Call SetTextboxPopupIfExist("TCNT-PROF_DB","Database prof\.","000000000001")
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call Settextbox("Project","CN_PROJN-LOW","",DT_CJ30_1000_PROJECT,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call ActivateItemGuiTree(0,"Object","Object")
Call TakeScreenShot

'''Call VerifyNodeTextGuiTree(2,DT_CJ30_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1)
'SAPGuiSession("Session_2").SAPGuiWindow("Execute Drilldown Report").SAPGuiTree("TableTreeControl_2").CheckProperty 
'''SAPGuiSession("Session_2").SAPGuiWindow("Execute Drilldown Report").SAPGuiTree("TableTreeControl_2").ActivateItem 
'''SAPGuiSession("Session_2").SAPGuiWindow("Execute Drilldown Report").SAPGuiTree("TableTreeControl_2").ActivateItem "Result","5.000"
'''SAPGuiSession("Session_2").SAPGuiWindow("Execute Drilldown Report").SAPGuiTree("TableTreeControl_2").ActivateItem "Result","5.000"

''' SelectItemGuiTree(treeIndex, itemPath, itemText)
'''Call SelectItemGuiTree(1,"Result","5.000")
Call SelectItemGuiTree(1,DT_CJ30_ITEM_PATH,DT_CJ30_ITEM_TEXT)


Call LogOff()
Call FinalStatus()

