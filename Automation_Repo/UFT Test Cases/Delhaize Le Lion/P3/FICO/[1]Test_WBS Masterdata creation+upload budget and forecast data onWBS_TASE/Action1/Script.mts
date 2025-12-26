
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
gstrTestCaseName = "Test_WBS Masterdata creat"
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

Call SetTextbox("Controlling Area","SVALD-VALUE","",DT_CJR2_0300_CONTROLLING_AREA,True)
Call ClickButton("Continue   \(Enter\)",True)
Call SetTextbox("Version","KPP0B-VALUE","",DT_CJR2_1000_VERSION,False)
Call SetTextbox("Fiscal Year","KPP0B-VALUE","",DT_CJR2_1000_FISCAL_YEAR,False)
Call SelectMenuBar("Extras;Excel Planning;Upload...")
Call SelectRadioButton("P_F_D_D","Import file directory",False)'''DT_CJR2_1000_IMPORT_FILE_DIRECTORY''''
Call SetTextbox("Path or file","FILENAME","",DT_CJR2_1000_PATH_OR_FILE,False)
Call FocusTextBox("File descriptions","FILE_ID",False)
Call SendKey("{F4}")
Call ClickLabel("01_1P7_ZDWBSXL0_001Z1-701","0",True)
Call SendKey("{F2}")
Call ClickButton("Execute   \(F8\)",False)
' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
Call VerifyifGuiLabelExistsByRelativeid(DT_CJR2_0120_CHECK_TEXT_OF_OK,"wnd\[0\]/usr/lbl\[78,1\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_CJR2_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[109,1\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_CJR2_0120_CHECK_TOOLTIP_OF_NO_NAME,"wnd\[0\]/usr/lbl\[109,1\]")

Call SetTcode(DT_CJR2_0120_OKCD)     
Call PressEnter() 
Call TakeScreenShot
Call SendKey("{F4}")
Call TakeScreenShot
Call SendKey("{F2}")
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Project","CN_PROJN-LOW","",DT_CJR2_1000_PROJECT,False)
Call FocusTextBox("Plan version","PAR_77",False)
Call TakeScreenShot
Call SendKey("{F4}")
Call Clicklabel("F1","0",True)
Call TakeScreenShot
Call SendKey("{F2}")
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
'''Call VerifyNodeTextGuiTree("0",DT_CJR2_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1)

'''Call SelectItemGuiTree(1,"Result","141.508-")
Call SelectItemGuiTree(1,DT_CJR2_ITEM_PATH,DT_CJR2_ITEM_TEXT)



'Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call SetTextbox("Project def\.","\*PROJ-PSPID","",DT_CJ01_0100_PROJECT_DEF,False)
'' SetComboByKey(attachedTextOrComboName, keyValue)
'Call SetComboByKey("Project Profile",DT_CJ01_0100_PROJECT_PROFILE)
'Call TakeScreenShot
'Call ClickButton("Project definition   \(Shift\+F1\)",False)
'Call TakeScreenShot
'Call SetTextbox("Company code","PROJ-VBUKR","",DT_CJ01_1205_COMPANY_CODE,False)
'Call SetTextbox("Profit Center","PROJ-PRCTR","",DT_CJ01_1205_PROFIT_CENTER,False)
'Call SetTextbox("Project def\.","PROJ-POST1","",DT_CJ01_0998_PROJECT_DEF,False)
'Call TakeScreenShot
'Call ClickButton("Save   \(Ctrl\+S\)",False)
'Call VerifyStatusBar(DT_CJ01_0100_CHECK_TEXT_OF_STATUSBAR)
'
Call LogOff()
Call FinalStatus()

