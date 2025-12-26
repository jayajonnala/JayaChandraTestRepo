'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Depreciate Assets_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th March
'.................Modified By :
'.................Modified Date/Details :
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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Depreciate Assets_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Sales_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Company Code","P_BUKRS","",DT_AFAB_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal year","P_GJAHR","",DT_AFAB_1000_FISCAL_YEAR,False)
Call SetTextbox("Main asset number","S_ANLN1-LOW","",DT_AFAB_1000_MAIN_ASSET_NUMBER,False)
Call SetTextbox("Posting Period","P_BUPER","",DT_AFAB_1000_POSTING_PERIOD,False)

Call SelectRadioButton("P_AGAIN","Repeat",False)
Call SelectCheckbox("P_TEST","Test Run","OFF",False)
Call TakeScreenShot()
Call SelectMenuBar("Program;Execute in Background")
Call TakeScreenShot()
Call PressEnter()     ' 
Call SetTextbox("Output Device","PRI_PARAMS-PDEST","",DT_OUTPUT_DEVICE,False)
Call SetCombo("Print Time","Immediately")
wait 5
'''''Call ClickButton("Continue   \(Shift\+F1\)",True)
Call ClickButton("Continue   \(Enter\)",True)
wait 5
Call ClickButton("Continue   \(Shift\+F1\)",True)


Call ClickButton("Immediate",True)
Call ClickButton("Save   \(Ctrl\+S\)",True)
Call TakeScreenShot()
Call GetStatusBar("text","DT_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AFAB_1000_CHECK_TEXT_OF_STATUSBAR)
'
''''''''''''''''''''SMX''''''''''''''''

Call SetTcode(DT_SAPTCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call ClickButton("Sort in Descending Order   \(Ctrl\+Shift\+F4\)",False)
Call ClickButton("Hide all fields \(CTRL\+F1\)",True)
Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST_820","Content","Start date",False)
Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call ClickLabel("B_RKHI",0,False)
Call TakeScreenShot()
Call ClickLabel("B_ICON",0,False)
Call TakeScreenShot()

''''''''''''''''AW01N''''''''''''''''''
Call SetTcode(DT_SAPTCODE_OCC1) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call SetTextbox("Asset","ANLA-ANLN1","",DT_ASSET,False)

Call PressEnter()   
wait 5
Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")
Call SelectTab("IDC_TABSTRIP","Posted values",False)
Call TakeScreenShot()
Call SetTextbox("Fiscal year","EDIT_JAHRE","",DT_YEAR,False)
Call PressEnter()   
Call TakeScreenShot()
Call DoubleClickGuiGridCell("Depreciation posted/planned",0,1,"Depreciation period",False)
Call TakeScreenShot()
wait 5

Call DoubleClickGuiGridCell("",0,1,"Reference document",False)
wait 3
Call DoubleClickGuiGridCell("Documents in Accounting",0,1,"Object type text",True)
Call TakeScreenShot()
Call GetTextBoxValue("BKPF-BELNR",0,"DT_DOCUMENT_OUTPUT",False)



Call LogOff()
Call FinalStatus ()




