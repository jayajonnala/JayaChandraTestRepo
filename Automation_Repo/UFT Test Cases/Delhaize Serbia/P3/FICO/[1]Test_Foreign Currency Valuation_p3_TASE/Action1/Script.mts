'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
 '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Foreign Currency Valuation_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 17th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Foreign Currency Valuation_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Foreign Currency Valuation_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 
'Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_PERIOD")
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''----------------------Tcode FBL3N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FBL3N_1000_COMPANY_CODE,False)

Call ActivateNodeGuiTree("","GL A/C Master Record;Account Group")
Call ActivateNodeGuiTree("","GL A/C Master Record;Trading partner")
Call ActivateNodeGuiTree("","Company code;Planning group")
Call ActivateNodeGuiTree("","Company code;Planning Level")
Call ActivateNodeGuiTree("","Company code;Account currency")
Call ActivateNodeGuiTree("","Document;Document Number")
'Capture the screenshot
Call TakeScreenShot()

wait(1)
SapGuiSession("transaction:=FBL3N").SapGuiWindow("transaction:=FBL3N").SAPGuiButton("tooltip:=Multiple selection","index:=5").Click
'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_FBL3N_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_FBL3N_3010_TABLECELL_SINGLE_VALUE_1,True)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","%%DYN004-LOW","",Replace((DT_DATE),"/","."),False)

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call SelectMenuBar("Edit;Find")
Call SetTextbox("Find","RSYSF-STRING","",DT_FBL3N_0800_FIND,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Find   \(Enter\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME)
Call ClickLabel(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME, "", True)
Wait(2)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)

'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Edit;Find")
Call SetTextbox("Find","RSYSF-STRING","",DT_FBL3N_0800_FIND_OCC1,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Find   \(Enter\)",True)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call ClickLabel(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC1, "", True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Cancel   \(F12\)",False)

'''----------------------Tcode TCURMNT----------------------------
'Enter the Tcode
Call SetTcode(DT_FBL3N_0120_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FBL3N_0120_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRowGuiTable("SAPMCURRTC_WORKSETS","Worklist",DT_FBL3N_0100_SAPMCURRTC_WORKSETS,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Exchange Rates   \(F5\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'Verify Table contents
Call VerifyTableCellContent(3,"Rate","SAPMCURRTC_RATE_SHEET",DT_FBL3N_0201_CHECK_TEXT_OF_TABLECELL_RATE_2)
Call VerifyTableCellContent(3,"Relation","SAPMCURRTC_RATE_SHEET",DT_FBL3N_0201_CHECK_TEXT_OF_TABLECELL_RELATION_2)
Call VerifyTableCellContent(5,"Rate","SAPMCURRTC_RATE_SHEET",DT_FBL3N_0201_CHECK_TEXT_OF_TABLECELL_RATE_4)
Call VerifyTableCellContent(5,"Relation","SAPMCURRTC_RATE_SHEET",DT_FBL3N_0201_CHECK_TEXT_OF_TABLECELL_RELATION_4)
''
'''----------------------Tcode FAGL_FCV----------------------------
''Enter the Tcode
Call SetTcode(DT_FBL3N_0201_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FBL3N_0201_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Get Variant\.\.\.   \(Shift\+F5\)",False)

Call SetTextbox("Variant","V-LOW","","",True)
Call SetTextbox("Environment","ENVIR-LOW","","",True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call SetTextbox("Changed By","AENAME-LOW","","",True)
Call SetTextbox("Original Language","MLANGU-LOW","","",True)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",True)

'select variant of row no 3
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program FAGL_FCV","",DT_FBL3N_0600_GRIDCELL_2_VARIANT_NAME,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(F2\)",True)

'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Valuation Key Date","P_VDATE","",Replace((DT_FBL3N_1000_VALUATION_KEY_DATE1),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()
'
Call SelectTab("TABSTRIP_SCRN_TAB",DT_FBL3N_1000_OPEN_ITEMS_SUBLEDGER,False)
'Capture the screenshot
Call TakeScreenShot()

SapGuiSession("transaction:=FAGL_FCV").SapGuiWindow("transaction:=FAGL_FCV").SAPGuiButton("tooltip:=Multiple selection","index:=4").Click
'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_FBL3N_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_FBL3N_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP_SCRN_TAB",DT_FBL3N_1000_OPEN_ITEMS_GL_ACCOUNTS,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP_SCRN_TAB",DT_FBL3N_1000_GL_ACCOUNT_BALANCES,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP_SCRN_TAB",DT_FBL3N_1000_OUTPUT__TECHNICAL_SETTINGS,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP_SCRN_TAB",DT_FBL3N_1000_POSTINGS,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(20)
'Capture the screenshot
Call TakeScreenShot()

'bellow step is for button "184 Postings   \(Shift\+F6\)" as the initial nummber is variable
Call SendKey("+{F6}")
'Capture the screenshot
Call TakeScreenShot()

'set filter by Doc header text
Call ClickLabel("Document Header Text", "", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Multiple selection",True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,"Reverse posting",True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,"FC Valuation",True)
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(Enter\)",True)
'set filter by posting key 50
Call ClickLabel("PK", "", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Execute   \(Enter\)",True)
Call SetTextbox("Posting Key","%%DYN001-LOW","","50",True)
Call ClickButton("Execute   \(Enter\)",True)
'verify labels
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC5)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC11)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC15)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC17)
'Capture the screenshot
Call TakeScreenShot()

'Change filter by posting key 40
Call ClickLabel("PK", "", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Execute   \(Enter\)",True)
Call SetTextbox("Posting Key","%%DYN001-LOW","","40",True)
Call ClickButton("Execute   \(Enter\)",True)
'verify labels
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC7)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC9)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC13)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC19)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)

Call SelectRadioButton("P_PPOST",DT_FBL3N_1010_POST_VALUATION_IMMEDIATELY,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(20)
'Capture the screenshot
Call TakeScreenShot()

'bellow step is for button "4 Message   \(Shift\+F8\)" as the initial nummber is variable
Call SendKey("+{F8}")
'Capture the screenshot
Call TakeScreenShot()

'get details and verify

Call VerifyGridCellContent("",2,"Message Type","",DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_0__ICON)
Call GetGridContentByTitle("","","Message Text",2,"DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_0_T_MSG_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("",2,"Message Text","",DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_0_T_MSG_OCC1)

Call VerifyGridCellContent("",3,"Message Type","",DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_1__ICON)
Call GetGridContentByTitle("","","Message Text",3,"DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_1_T_MSG_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("",3,"Message Text","",DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_1_T_MSG_OCC1)

Call VerifyGridCellContent("",4,"Message Type","",DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_2__ICON)
Call GetGridContentByTitle("","","Message Text",4,"DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_2_T_MSG_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("",4,"Message Text","",DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_2_T_MSG_OCC1)

Call VerifyGridCellContent("",5,"Message Type","",DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_3__ICON)
Call GetGridContentByTitle("","","Message Text",5,"DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_3_T_MSG_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("",5,"Message Text","",DT_FBL3N_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_3_T_MSG_OCC1)

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

'bellow step is for button "184 Postings   \(Shift\+F6\)" as the initial nummber is variable
Call SendKey("+{F6}")
'Capture the screenshot
Call TakeScreenShot()

'set filter by Doc header text
Call ClickLabel("Document Header Text", "", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Multiple selection",True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,"Reverse posting",True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,"FC Valuation",True)
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(Enter\)",True)
'verify labels of newly generated docs
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC20)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC23)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC26)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC29)

'set filter by posting key 50
Call ClickLabel("PK", "", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Execute   \(Enter\)",True)
Call SetTextbox("Posting Key","%%DYN001-LOW","","50",True)
Call ClickButton("Execute   \(Enter\)",True)
'verify labels
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC21)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC25)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC28)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC30)

'Capture the screenshot
Call TakeScreenShot()

'Change filter by posting key 40
Call ClickLabel("PK", "", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Execute   \(Enter\)",True)
Call SetTextbox("Posting Key","%%DYN001-LOW","","40",True)
Call ClickButton("Execute   \(Enter\)",True)
'verify labels
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC22)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC24)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC27)
Call VerifyifGuiLabelExists(DT_FBL3N_0120_CHECK_TEXT_OF_NO_NAME_OCC31)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

