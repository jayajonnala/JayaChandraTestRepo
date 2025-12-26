
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_069-Maintain Vendor Article-Code and check MII033_P2_WE09 ZMDPU_VEND_ART
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

gstrTestCaseName = "Test_MD_01_01_069_P2_WE09 ZMDPU_VEND_ART"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


'''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''--------TransactionCode-WE09----------''''

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
'Call ClickButton("Sort in descending order   \(Ctrl\+Shift\+F4\)",False)
Call ClickButton("Sort in Ascending Order   \(Ctrl\+F4\)",False)
 
Call TakeScreenShot()

Call GetLabelContentByRefLabel("IDoc number",0,-48,"DT_IDOC_OUTPUT",False)
'Call GetLabelContentByRefLabel("IDoc number",0,-72,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call TakeScreenShot()
'
''Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;Segment 000001;Segment 000004","ZZMDPU_VEND_ART_I","",False)
Call ClickLinkGuiTree("#1;#2;#1;#2","#1",0,False)


Call TakeScreenShot()

Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", "", DT_WE09_0100_CHECK_TEXT_OF_CURRENT_STATUS, False)

Call VerifyTableCellContent(1, "Fld Cont.", "IDOC_TREE_CONTROLINT_SEG_CONTROL", DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0)
Call VerifyTableCellContent(2, "Fld Cont.", "IDOC_TREE_CONTROLINT_SEG_CONTROL", DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_1)

'''--------TransactionCode----/nzmdpu_vend_art----------''''

Call SetTcode(DT_WE09_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Purch. Organization","P_EKORG","",DT_WE09_1000_PURCH_ORGANIZATION,False)
'Call SetTextbox("Vendor","P_LIFNR","",DT_WE09_1000_VENDOR,False)
Call SetTextboxNoLabel("P_LIFNR","",DT_WE09_1000_VENDOR,False)
Call SetTextbox("Article","P_MATNR","",DT_WE09_1000_ARTICLE,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()


Call ClickButton("Change\/Display   \(F7\)",False)


'''Selectcell----Add 2 Articles'''
Wait 2
SAPGuiSession("Session").SAPGuiWindow("Creation of multiple vendor's").InsightObject("InsightObject").Click

Wait 2
SAPGuiSession("Session").SAPGuiWindow("Creation of multiple vendor's").InsightObject("InsightObject").Click

Wait 2



Call SetGridData("",2,"Vendor Article Number",DT_WE09_0100_GRIDCELL_1_VENDOR_ARTICLE_NUMBER_OCC1,False)
Call SetGridData("",3,"Vendor Article Number",DT_WE09_0100_GRIDCELL_2_VENDOR_ARTICLE_NUMBER_OCC1,False)

Call SelectCheckBoxGridByRefColumn("shell","Primary Vendor Material Code - Index", "Vendor Article Number",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0, "OFF")
Call SelectCheckBoxGridByRefColumn("shell","Primary Vendor Material Code - Index", "Vendor Article Number",DT_WE09_0100_GRIDCELL_2_VENDOR_ARTICLE_NUMBER_OCC1, "ON")
 

Call ClickButtonToolBar("&DETAIL",0)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)

Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot

Call VerifyGridCellContentbyName("shell",3,"Vendor Article Number","",DT_WE09_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_2_IDNLF)



Call LogOff()

Call FinalStatus ()


