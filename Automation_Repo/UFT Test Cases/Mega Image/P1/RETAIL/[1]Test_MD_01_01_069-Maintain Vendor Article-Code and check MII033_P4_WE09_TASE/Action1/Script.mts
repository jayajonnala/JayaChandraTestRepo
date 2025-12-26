
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_069-Maintain Vendor Article-Code and check MII033_P4_WE09
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


gstrTestCaseName = "Test_MD_01_01_069-Maintain Vendor Article-Code and check MII033_P4_WE09"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''--------TransactionCode-WE09----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_WE09_1000_CREATED_ON),False)
Call SetTextbox("Logical Message","MESTYP-LOW","",DT_WE09_1000_LOGICAL_MESSAGE,False)

Call SetTextbox("Search in Field \.\.\.","FIELD1_1","",DT_WE09_1000_SEARCH_IN_FIELD_,False)
Call SetTextbox("for Value \.\.\.","VALUE1_1",0,DT_WE09_1000_FOR_VALUE_,False)
Call SetTextbox("and Search in Field \.\.\.","FIELD1_2","",DT_WE09_1000_AND_SEARCH_IN_FIELD_,False)
Call SetTextbox("for Value \.\.\.","VALUE1_2","",DT_WE09_1000_FOR_VALUE__OCC1,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call VerifyStatusBar("IDocs were found")
Call SetFocusGuiLabel("Time", "", "", False)
Call TakeScreenShot()

Call ClickLabel("Time",0,False)
Call ClickButton("Sort in descending order   \(Ctrl\+Shift\+F4\)",False)
'Call ClickButton("Sort in Ascending Order   \(Ctrl\+F4\)",False)
 
Call TakeScreenShot()

Call GetLabelContentByRefLabel("IDoc number",0,-48,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call TakeScreenShot()


Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDPU_VEND_ART_H","ZZMDPU_VEND_ART_H","",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", "", DT_WE09_0100_CHECK_TEXT_OF_CURRENT_STATUS, False)

'Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;Segment 000001;Segment 000003","ZZMDPU_VEND_ART_I","",False)
Call ClickLinkGuiTree("#1;#2;#1;#2","#1",0,False)

Call TakeScreenShot()
Call VerifyTableCellContent(1, "Fld Cont.", "IDOC_TREE_CONTROLINT_SEG_CONTROL", DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0)
Call VerifyTableCellContent(2, "Fld Cont.", "IDOC_TREE_CONTROLINT_SEG_CONTROL", DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_1)

Call LogOff()

Call FinalStatus ()



