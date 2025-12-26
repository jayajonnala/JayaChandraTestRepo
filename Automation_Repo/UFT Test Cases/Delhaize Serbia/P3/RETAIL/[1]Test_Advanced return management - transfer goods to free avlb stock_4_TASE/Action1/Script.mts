
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Advanced return management - transfer goods to free avlb stock_4
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

gstrTestCaseName = "Test_Advanced return management - transfer goods to free avlb stock_4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''''---------Login-------------''''''
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''-------Transaction Code MSR_INSPWH-------'''''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Delivery","SO_DLVNI-LOW","",DT_MSR_TRC_I_0301_DELIVERY,False)
Call TakeScreenShot
Call ClickButtonIFExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ActivateNodeGuiTree("","#2;#1")
Call TakeScreenShot
Call GetTextboxValue("GODYNPRO-MAT_DOC", "", "DT_MSR_TRC_I_0300_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MSR_TRC_I_0300_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT",DT_MSR_TRC_I_0300_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Back   \(F3\)",FalsE)
Call TakeScreenShot

'''''''-------Transaction Code MIGO-------'''''''

Call SetTcode(DT_MSR_TRC_I_0300_OKCD)     
Call PressEnter() 
Call TakeScreenShot

Call SetTextBoxNoLabel("GODYNPRO-MAT_DOC","",DT_MSR_TRC_I_2010_GODYNPROMAT_DOC,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call VerifyTableCellContent(1,"Movement Type","SAPLMIGOTV_GOITEM",DT_MSR_TRC_I_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)

Call SelectTab("TS_GOITEM","Messages",False)

Call ClickButton("Display outputs",False)

Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3", DT_MSR_TRC_I_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(1, "Status", "SAPDV70ATC_NAST3", DT_MSR_TRC_I_0100_CHECK_TEXT_OF_TABLECELL_STATUS_0)

Call VerifyTableCellContent(2, "Output Type", "SAPDV70ATC_NAST3", DT_MSR_TRC_I_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)
Call VerifyTableCellContent(2, "Status", "SAPDV70ATC_NAST3", DT_MSR_TRC_I_0100_CHECK_TEXT_OF_TABLECELL_STATUS_1)
'
Call LogOff'
Call FinalStatus()



