

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.2.1.2.2. Modify Vendor Account Group ZVXT - External Trade Ven_TASE
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.2.1.2.2. Modify Vendor Account Group ZVXT - External Trade Ven_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_2.2.1.2.2. Modify Vendor Account Group ZVXT - External Trade Ven_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------Tcode ZMDPU_INFOREC_COPY----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Call SetTextbox("Vendor","RF02K-LIFNR","",DT_XK02_0101_VENDOR,False)
Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK02_0101_VENDOR,False)
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK02_0101_COMPANY_CODE,False)
Call SetTextbox("Purchasing Organization","RF02K-EKORG","",DT_XK02_0101_PURCH_ORGANIZATION,False)

Call SelectCheckbox("RF02K-D0110","1","ON",False)
Call SelectCheckbox("RF02K-D0120","1","ON",False)
Call SelectCheckbox("RF02K-D0130","1","ON",False)
Call SelectCheckbox("RF02K-D0210","1","ON",False)
Call SelectCheckbox("RF02K-D0215","1","ON",False)
Call SelectCheckbox("RF02K-D0220","1","ON",False)
Call SelectCheckbox("RF02K-D0610","1","ON",False)
Call SelectCheckbox("RF02K-D0310","1","ON",False)
Call TakeScreenShot()
Call PressEnter()  
Call PressEnter()  


Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XK02_0301_STREETHOUSE_NUMBER,False)
Call SetTextbox("Mobile Phone","SZA1_D0100-MOB_NUMBER","",DT_XK02_0301_MOBILE_PHONE,False)


Call ClickButtonIfExist("Other E-Mail Addresses",False)
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("New Line   \(Shift\+F1\)",True)
Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","E-Mail Address",2,DT_XK02_0600_TABLECELL_EMAIL_ADDRESS_2,True)
Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","Notes",2,DT_XK02_0600_TABLECELL_NOTES_2,True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Next screen   \(F8\)",False)
Call ClickButtonIfExist("Next screen   \(F8\)",False)
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Bank Account",1,DT_XK02_0130_TABLECELL_BANK_ACCOUNT_0,False)
Call ClickButtonIfExist("Next screen   \(F8\)",False)
Call ClickButtonIfExist("Next screen   \(F8\)",False)
wait 5
Call SendKey("{F4}")

Call SetFocusGuiLabel(DT_XK02_LABEL_1,"11","232",True)
wait 5
Call SendKey("{F2}")
Call FocusTextBox("Pmnt Meth\. Sup\.","LFB1-UZAWE",False)
wait 5
Call SendKey("{F4}")
wait 5
Call SetVerticalScrollBar(10,True)
Call SetFocusGuiLabel("64","11","536",True)
Call SendKey("{F2}")
wait 2
Call ClickButtonIfExist("Next screen   \(F8\)",False)
Call ClickButtonIfExist("Next screen   \(F8\)",False)

Call FocusTextBox("Terms of paymnt","LFM1-ZTERM",False)
Call SendKey("{F4}")
wait 5
Call SetFocusGuiLabel(DT_XK02_LABEL_1,"11","232",True)
Call SendKey("{F2}")
wait 2
Call PressEnter() 
Call ClickButtonIfExist("Cancel",False)
Call ClickButton("Different Data   \(Shift\+F6\)",False)
Call SendKey("{F2}")
wait 2
Call SetTextbox("Planned deliv\. time","LFM2-PLIFZ","",DT_XK02_0330_PLANNED_DELIV_TIME,False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call VerifyStatusBar(DT_XK02_0101_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
