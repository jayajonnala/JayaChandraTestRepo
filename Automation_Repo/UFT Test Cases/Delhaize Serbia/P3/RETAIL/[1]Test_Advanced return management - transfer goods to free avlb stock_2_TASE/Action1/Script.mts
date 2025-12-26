
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Advanced return management - transfer goods to free avlb stock_2
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_Advanced return management - transfer goods to free avlb stock_2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


'---------------------------Login----------------------------------
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'------------------------------Transaction Code ME23N---------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True)
Call ClickButtonifExist("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonifExist("Expand Header Ctrl\+F2",False)
Call TakeScreenShot
Call SelectTab("HEADER_DETAIL","Returns",False)
Call TakeScreenShot
Call ClickButton("Open Returns Overview",False)
Call TakeScreenShot

Call ActivateNodeGuiTree(0, "#1;#1")
Call TakeScreenShot
Call VerifyTextBoxNoLabelContent("MEPO_TOPLINE-EBELN", "", DT_ME23N_0003_PUR_ORDER,False)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call ActivateNodeGuiTree(0, "#1;#2")
Call TakeScreenShot
Call GetTextboxValue("LIKP-VBELN", "", "DT_ME23N_0200_CHECK_TEXT_OF_TREE_RETURNS_STOCK_TRANSPORT_ORDER_NODE_OCC1_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ME23N_0200_CHECK_TEXT_OF_TREE_RETURNS_STOCK_TRANSPORT_ORDER_NODE_OCC1_OUTPUT",DT_ME23N_0200_CHECK_TEXT_OF_TREE_RETURNS_STOCK_TRANSPORT_ORDER_NODE_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
'
Call ActivateNodeGuiTree(0, "#1;#3")
Call TakeScreenShot
Call GetTextboxValue("GODYNPRO-MAT_DOC", "", "DT_ME23N_0200_CHECK_TEXT_OF_TREE_RETURNS_STOCK_TRANSPORT_ORDER_NODE_OCC2_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ME23N_0200_CHECK_TEXT_OF_TREE_RETURNS_STOCK_TRANSPORT_ORDER_NODE_OCC2_OUTPUT",DT_ME23N_0200_CHECK_TEXT_OF_TREE_RETURNS_STOCK_TRANSPORT_ORDER_NODE_OCC2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call ActivateNodeGuiTree(0, "#1;#4")
Cll TakeScreenShot
Call GetTextboxValue("LIKP-VBELN", "", "DT_ME23N_0200_CHECK_TEXT_OF_TREE_RETURNS_STOCK_TRANSPORT_ORDER_NODE_OCC3_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ME23N_0200_CHECK_TEXT_OF_TREE_RETURNS_STOCK_TRANSPORT_ORDER_NODE_OCC3_OUTPUT",DT_ME23N_0200_CHECK_TEXT_OF_TREE_RETURNS_STOCK_TRANSPORT_ORDER_NODE_OCC3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

''''''''-------Transaction Code VL32N-------'''''''

Call SetTcode(DT_ME23N_0200_OKCD)     
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Inbound delivery","LIKP-VBELN","",DT_ME23N_4104_INBOUND_DELIVERY,False)
Call PressEnter() 
Call TakeScreenShot

Call ClickButton("Post Goods Receipt   \(Shift\+F8\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_ME23N_4104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME23N_4104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_ME23N_4104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call GetStatusBar("item2","DT_ME23N_4104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME23N_4104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT",DT_ME23N_4104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyStatusBar(DT_ME23N_4104_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus()


