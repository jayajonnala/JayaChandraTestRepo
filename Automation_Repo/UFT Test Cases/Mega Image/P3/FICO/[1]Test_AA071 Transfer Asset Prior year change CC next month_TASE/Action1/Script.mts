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


 '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_AA071 Transfer Asset Prior year change CC next month_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 21st June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_AA071 Transfer Asset Prior year change CC next month_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_AA071 Transfer Asset Prior year change CC next month_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

Call StartDateof445PeriodByDate(DT_TODAY,"DT_STARTING_DATE_PERIOD")
Call EndDateof445PeriodByDate(DT_TODAY,"DT_END_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'''Increment the parameter / reload
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_TEXT",(Cint(DT_INCREMENT_TEXT)+1))
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode AS02----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS02_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS02_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS02_0100_COMPANY_CODE,False)
Call TakeScreenShot

Call PressEnter() 
Call TakeScreenShot

Call SelectTab("TABSTRIP100","Time-dependent",False)
Call TakeScreenShot

Call ClickButtonIfExist("More Intervals",False)
Call TakeScreenShot

For Iterator = 1 To 10 Step 1
	If SAPGuiSession("transaction:=AS02").SapGuiWindow("transaction:=AS02").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'Call ClickButtonIfExist("Insert Row",False)
'Call TakeScreenShot
'
'For Iterator = 1 To 10 Step 1
'	If SAPGuiSession("transaction:=AS02").SapGuiWindow("transaction:=AS02").SAPGuiStatusBar("messagetype:=W").Exist = True Then
'		Call PressEnter() 
'	End If
'Next
'
'Call ClickButtonIfExist("Confirm   \(Enter\)",True)
'
'Call SetTextbox("From-date of new interval","ANLZ-ADATU","",DT_AS02_3010_FROMDATE_OF_NEW_INTERVAL,True)
'Call ClickButtonIfExist("Yes   \(Enter\)",True)
'Call TakeScreenShot

Call SetTableDataNoRef("SAPLAISTTIME","Cost Center",1,DT_AS02_3000_TABLECELL_COST_CENTER_0,False)
Call SetTableDataNoRef("SAPLAISTTIME","Profit Center",1,DT_AS02_3000_TABLECELL_PROFIT_CENTER_0,False)

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)

Call GetTextStatusBar("DT_AS02_3000_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

For Iterator = 1 To 10 Step 1
	If SAPGuiSession("transaction:=AS02").SapGuiWindow("transaction:=AS02").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next 
Wait(2)
Call GetStatusBar("item1","DT_AS02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_SUCCESSMESSAGE)
Call TakeScreenShot

'
'''----------------------Tcode S_ALR_87011966----------------------------
'Enter the Tcode
Call SetTcode(DT_AS02_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS02_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("All Selections   \(Shift\+F7\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("XEINZEL",DT_AS02_1000_LIST_ASSETS,False)
Call SetTextbox("Company code","BUKRS-LOW","",DT_AS02_0100_COMPANY_CODE,False)
Call SetTextbox("Asset number","ANLAGE-LOW","",DT_AS02_1000_ASSET_NUMBER,False)
Call SetTextbox("Subnumber","UNTNR-LOW","",DT_AS02_1000_SUBNUMBER,False)
Call SetTextbox("Cost center","SO_KOSTL-LOW","",DT_AS02_1000_COST_CENTER,False)
Call SetTextbox("Report date","BERDATUM","",Replace((DT_AS02_1000_REPORT_DATE),"/","."),False)
Call SetTextbox("Sort Variant","SRTVR","",DT_AS02_1000_SORT_VARIANT,False)
Call SetTextbox("Display variant","P_VARI","",DT_AS02_1000_DISPLAY_VARIANT,False)
Call PressEnter() 
Call TakeScreenShot
 
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Information Message","MESSTXT1",0,Lcase(DT_AS02_0010_CHECK_TEXT_OF_MESSTXT1_OCC1),True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Cost center","SO_KOSTL-LOW","","",False)'DT_AS02_1000_COST_CENTER_OCC1
Call SetTextbox("Report date","BERDATUM","",Replace((DT_AS02_1000_REPORT_DATE_02),"/","."),False)
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Wait(3)

Call GetGridContentByTitle("","","Asset",1,"DT_AS02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN0_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("",1,"Asset",0,DT_AS02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN0_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


