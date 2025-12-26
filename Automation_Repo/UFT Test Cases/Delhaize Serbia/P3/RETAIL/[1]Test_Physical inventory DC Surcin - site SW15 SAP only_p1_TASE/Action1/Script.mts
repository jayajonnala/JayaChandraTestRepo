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
'.................Test Script Name : Test_Physical inventory DC Surcin - site SW15 SAP only_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 7th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Physical inventory DC Surcin - site SW15 SAP only_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Physical inventory DC Surcin - site SW15 SAP only_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MI31_0731_ISEGERFMG",(Cint(DT_MI31_0731_ISEGERFMG)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode MI31----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Article","R_MATNR-LOW",False)
Call ClickButton("%_R_MATNR_%_APP_%-VALU_PUSH",False)

''wait(2)
''Call SendKey("{TAB}")
''wait(2)
''Call SendKey("{TAB}")
''wait(2)
''Call SendKey("{ENTER}")
''wait(2)

Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_ARTICLE,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_ARTICLE_2,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_ARTICLE_3,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",4,DT_ARTICLE_4,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call SelectCheckbox("XDELE","1","ON",False)
Call SetTextbox("Site","R_WERKS-LOW","",DT_MI31_1000_SITE,False)
''Call SetTextbox("Storage Location","R_LGORT-LOW","",DT_MI31_1000_STORAGE_LOCATION,False)
Call SetTextbox("Storage Location","R_LGORT-LOW","","",False)
Call TakeScreenShot()

Call SelectCheckbox("XKEEP","1",DT_MI31_1000_HOLD_PROCESSED_SESSIONS,False)
Call SetTextbox("Name of Session","MAPPE","",DT_MI31_1000_NAME_OF_SESSION,False)
Call SetTextbox("Max\. No\. Items/Doc\.","MAXPO","",DT_MI31_1000_MAX_NO_ITEMSDOC,False)
Call TakeScreenShot()

Call FocusTextBox("Max\. No\. Items/Doc\.","MAXPO",False)
Call ClickButtonIfExist("Acc\. to Stck",False)
Wait(2)
Call SelectCheckbox("XIMAT","1",DT_MI31_1000_INCL_ARTLS_SUBJ_TO_PHYS_INV,False)
Call SelectCheckbox("XLABST","1",DT_MI31_1000_UNRESTRICTED_USE,False)
Call SelectCheckbox("XBUFI","1",DT_MI31_1000_FREEZE_BOOK_INVBAL,False)
Call SetTextbox("Planned Count Date","GIDAT","",Replace((DT_MI31_1000_PLANNED_COUNT_DATE),"/","."),False)
Call TakeScreenShot()

Call SetTextbox("Physical Inventory Number","INVNU","",DT_MI31_1000_PHYSICAL_INVENTORY_NUMBER,False)
Call FocusTextBox("Physical Inventory Number","INVNU",False)
Call TakeScreenShot()
Call PressEnter() 
Call GetTextboxValue("INVNU","","DT_MI31_1000_CHECK_TEXT_OF_PHYSICAL_INVENTORY_NUMBER_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_MI31_1000_CHECK_TEXT_OF_PHYSICAL_INVENTORY_NUMBER_OUTPUT",DT_MI31_1000_CHECK_TEXT_OF_PHYSICAL_INVENTORY_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)
Wait(2)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call TakeScreenShot()

Call ClickButtonIfExist("Process Session   \(Shift\+F2\)",False)
Call TakeScreenShot()

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
Call ClickButtonIfExist("Process session   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call SelectCheckbox("D0300-DYDEFSIZE","1",DT_MI31_0300_DEFAULT_DYNPRO_SIZE,True)
Call SelectRadioButton("D0300-BATCH",DT_MI31_0300_BACKGROUND,True)
Call TakeScreenShot()

Call ClickButtonIfExist("Process   \(Enter\)",True)
Wait(2)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
Call TakeScreenShot()
'
'''----------------------Tcode MI01----------------------------
'Enter the Tcode
Call SetTcode(DT_MI31_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MI31_1000_OKCD)
Call TakeScreenShot()

Call SetTextbox("Count Date","RM07I-ZLDAT","",Replace((DT_MI31_0701_COUNT_DATE),"/","."),False)
Call TakeScreenShot()
Call SendKey("{F4}")
Wait(3)

Call SetTextbox("Article","IM_MATNR-LOW","","",False)
Call SetTextbox("Site","IM_WERKS-LOW","","",False)
Call SetTextbox("Storage Location","IM_LGORT-LOW","","",False)
Call SetTextbox("Physical Inventory Number","IM_INVNU-LOW","",DT_MI31_1000_PHYSICAL_INVENTORY_NUMBER_OCC1,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
Wait(2)
Call GetLabelContentByRefLabel("PhysInvDoc","0","-48","DT_SELECT_DOC_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''SAPGuiSession("transaction:=MI22").SAPGuiWindow("transaction:=MI22").SAPGuiLabel("guicomponenttype:=30","content:="&DT_SELECT_DOC,"index:=0").SetFocus

Call SetFocusGuiLabel("","11","88",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Adopt Phys\. Inv\. Doc\. No\.   \(Ctrl\+Shift\+F9\)",False)
Wait(2)
Call TakeScreenShot()
Call PressEnter()  

Call SetTextboxNoLabel("ISEG-ERFMG","",DT_MI31_0731_ISEGERFMG,False)
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(3)
Call VerifyStatusBarMessageType("S")
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_MI31_0701_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_MI31_0701_CHECK_TEXT_OF_STATUSBAR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

