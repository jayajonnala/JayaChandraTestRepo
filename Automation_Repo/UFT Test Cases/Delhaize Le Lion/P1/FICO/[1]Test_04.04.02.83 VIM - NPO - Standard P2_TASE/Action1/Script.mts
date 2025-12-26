
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.83 VIM - NPO - Standard P2
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

gstrTestCaseName = "Test_04.04.02.83 VIM - NPO - Standard P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''all LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''

Call SetTcode(DT_OPTVIM_7AX2_0002_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

'--------------------------------------------  /OPT/VIM_WP----------------------------------------------



Call SetTextbox("Document Id","H1_DOID-LOW","",DT_OPTVIM_7AX2_1105_DOCUMENT_ID,False)
Call ClickButton("BT_H_APPLY",False)
Call ClickButton("Switch Work View   \(Shift\+F6\)",False)
Call TakeScreenShot()
Call SelectRadioButton("SPOPLI-SELFLAG","All Users View",True)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Hide Detail Pane   \(Ctrl\+F2\)",False)

Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False)
Call SelectTab("TAB_MAIN","Process",False)
Call SelectTab("TAB_MAIN","Basic Data",False)
Call SetTextboxNoLabel("GH_IDX_APPLICATION->MS_IDX_HEADER-ZZKOSTL",0,FormatBlank(DT_OPTVIM_7AX2_9100_GH_IDX_APPLICATIONMS_IDX_HEADERZZKOSTL),False)
Call PressEnter()  
Call PressEnter()  
Call SelectTab("TAB_MAIN","Accounting",False)
Call PressEnter()  
Call PressEnter() 
'''''Steps swapped(90-92 <-> 68-73) based on Defect 20277
Call SelectTab("TAB_MAIN","Line Items",False)
Call TakeScreenShot()
Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","G/L",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_GL_ACC_0,False)
Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Amount",1,"","",DT_OPTVIM_7AX2_0019_TABLECELL_AMOUNT_0,False)
Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Cost Center",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_COST_CTR_0,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Show Options   \(Ctrl\+F12\)",False)
Call ClickCellGuiGrid("",0,"Option",2,"","",False)
''Call ClickCellGuiGrid("",0,"Option Short Text",2,"","",False)
 Wait 25
Call ClickButton("BT_H_APPLY",False)
wait 5
Call ClickButton("Refresh   \(F5\)",False)
'**************************** Updated below script - based on the Zephyr defect 30608- inputs for Potential rent invoice isue*************
Call TakeScreenShot()
Call GetGridContentByTitle("All Inbox.*",0,"Exception Reason",1,"DT_OPTVIM_7AX2_2004_TABLECELL_EX_REA_OUTPUT")
If DT_OPTVIM_7AX2_2004_TABLECELL_EX_REA_OUTPUT =  DT_OPTVIM_7AX2_2004_TABLECELL_EX_REA Then
	Call TakeScreenShot()
	Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
	Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
	Call TakeScreenShot()
	'ClickCellGuiGrid(gridTitle,gridIndex,columnName,rowNumber,columnNameRef,gridValRef,blnIsItPopup)
	Call ClickCellGuiGrid( "","","Activate/Bypass","9","Business Rule","Potential Rent Invoice (NPO)",False)
	Call TakeScreenShot()
	Call ClickButtonIfExist("Yes",True)
	Call SetTextArea(DT_OPTVIM_7AX2_2004_BYPASS_COMMENTS)
	Call TakeScreenShot()
	Call ClickButton("Save   \(F2\)",True)
	Call ClickButton("Exit   \(Shift\+F7\)",True)
	Call TakeScreenShot()
	Call ClickCellGuiGrid("",0,"Option",2,"","",False)
	 Wait 25
	Call ClickButton("BT_H_APPLY",False)
	wait 5
	Call ClickButton("Refresh   \(F5\)",False)
	Call TakeScreenShot()
	Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
	Call TakeScreenShot()
End If
'*********************************************************
while len(DT_OPTVIM_7AX2_1105_DOCUMENT_ID_OLD)<12 
	DT_OPTVIM_7AX2_1105_DOCUMENT_ID_OLD="0"&DT_OPTVIM_7AX2_1105_DOCUMENT_ID_OLD
Wend
Call VerifyWindowValue("Approve Invoice "&DT_OPTVIM_7AX2_1105_DOCUMENT_ID_OLD&" on behalf of VIM_CODER_BE  VIM_CODER_BE")
'''Call VerifyWindowValue(DT_OPT_VIM_CHECK_1ST_APPROVER)
Call TakeScreenShot()

''Call SetTableData("/ORS/SAPL000007_APPR_SCREENTC_LINE","G/L Acc",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_GL_ACC_0,False)
''Call SetTableData("/ORS/SAPL000007_APPR_SCREENTC_LINE","Cost Ctr",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_COST_CTR_0,False)
''Call PressEnter() 
Call ClickButton("Approve",False)
Call ClickButton("APPR_BUTTON",False)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
wait 25
Call ClickButton("Refresh   \(F5\)",False)
Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
Call VerifyWindowValue("Approve Invoice "&DT_OPTVIM_7AX2_1105_DOCUMENT_ID_OLD&" on behalf of   ASRMDLL01")
''Call VerifyWindowValue(DT_OPT_VIM_CHECK_2ND_APPROVER)
Call TakeScreenShot()

Call ClickButton("Approve",False)
Call ClickButton("APPR_BUTTON",False)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
Call VerifyWindowValue("Approve Invoice "&DT_OPTVIM_7AX2_1105_DOCUMENT_ID_OLD&" on behalf of   ASRMDLL03")
''Call VerifyWindowValue(DT_OPT_VIM_CHECK_3RD_APPROVER)
Call TakeScreenShot()

Call ClickButton("Approve",False)
Call ClickButton("APPR_BUTTON",False)
Call ClickButton("Back   \(F3\)",False)

'--------------------------------------------   /n/opt/vim_va2----------------------------------------------
Call SetTcode(DT_OPTVIM_7AX2_1000_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Document Processing Number","S_DOCID-LOW","",DT_OPTVIM_7AX2_1000_DOCUMENT_PROCESSING_NUMBER,False)
Call ClickButton("Execute   \(F8\)",False)
wait 10
Call ClickButton("Refresh   \(F5\)",False)
Call VerifyGridCellContent("Results \(1 Hit\)",1,"OVERALL_STATUS_TEXT",0,DT_OPTVIM_7AX2_2000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_OVERALL_STATUS_TEXT)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()
