
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.46 VIM - PO Precontrole Issue - BR24 - Missing Posting Date (PO)
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


gstrTestCaseName = "Test_Retrive child and post GI + Verifications"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''''Caall LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'--------------------------------------------  SE16----------------------------------------------

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call TakeScreenShot()

Call VerifyTableCellContent(1,"Article", ".*", DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(2,"Article", ".*", DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call VerifyTableCellContent(1,"Deliv. Qty", ".*" ,DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_0)
'Call VerifyTableCellContent(2,"Deliv. Qty", ".*", DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_1)
Call VerifyTableCellContent(1,"Site", ".*", DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SITE_0)
Call VerifyTableCellContent(2,"Site", ".*", DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SITE_1)
Call VerifyTableCellContent(1,"SLoc", ".*",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SLOC_0)
Call VerifyTableCellContent(2,"SLoc", ".*", DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SLOC_1)
Call VerifyTextBoxContent("Ship-To Party","KUWEV-KUNNR",0,DT_VL03N_1502_CHECK_TEXT_OF_SHIPTO_PARTY,False)
Call TakeScreenShot()

Call ClickButton("Header Details   \(F8\)",false)
Call SelectTab("TAXI_TABSTRIP_HEAD","Administration",False)
Call TakeScreenShot()
Call GetTextboxValue("LIKP-LIFEX",0,"DT_VL03N_2110_CHECK_TEXT_OF_EXT_DELIVERY_OUTPUT",False)
Call TakeScreenshot()
Call ClickButtonIfExist("Back   \(F3\)",False)
wait 5
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButton("Display <-> Change   \(Ctrl\+F1\)",false)
Call TakeScreenShot()
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",false)

'If VerifyStatusBarConditon(DT_VL03N_4004_CHECK_TEXT_OF_STATUSBAR)= True then
If VerifyStatusBarConditonNoReport(DT_VL03N_4004_CHECK_TEXT_OF_STATUSBAR) <> True then
     Call VerifyGridCellContent("",1,"Message Text",0,DT_VL03N_4004_CHECK_TEXT_OF_STATUSBAR_OCC1)
       Call TakeScreenShot  
Else
       Call VerifyStatusBar(DT_VL03N_4004_CHECK_TEXT_OF_STATUSBAR)
       Call TakeScreenShot
End If

''''''''''''---------------------------vl03n----------''''''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY,FALSE)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Document Flow   \(F7\)",False)
Wait(2)
Call TakeScreenShot()

If ValidateNodeTextGuiTree(0,"GD goods issue:") = True Then
	Call ActivateNodeGuiTree(0,"#1;#1;RegExp:=GD.*")
	Call GetGridContent("GD.*", 0, "Doc.no.", 1, "Item","", "DT_VL03N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")
	Call TakeScreenShot()

Else
	Call ActivateNodeGuiTree(0,"#1;#1;RegExp:=TF.*")
	Call GetGridContent("TF.*", 0, "Doc.no.", 1, "Item","", "DT_VL03N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")
       Call TakeScreenShot()	
End If

Call LogOff()
Call FinalStatus ()


