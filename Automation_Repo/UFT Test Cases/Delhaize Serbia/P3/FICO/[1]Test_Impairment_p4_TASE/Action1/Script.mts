'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Impairment_p4_TASE
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

gstrTestCaseName = "Test_Impairment_p4_TASE"
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
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Company Code","P_BUKRS","",DT_ABAW_0100_COMPANY_CODE,False)
'Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABAW_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal year","P_GJAHR","",DT_YEAR,False)
Call SetTextbox("Main asset number","S_ANLN1-LOW","",DT_ABAW_0100_ASSET,False)
Call SetTextbox("Posting Period","P_BUPER","",DT_ABAW_0100_POSTING_PERIOD,False)

Call SelectRadioButton("P_AGAIN","Repeat",False)
Call SelectCheckbox("P_TEST","Test Run","OFF",False)
Call TakeScreenShot()
Call SelectMenuBar("Program;Execute in Background")
Call TakeScreenShot()
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
VerifyStatusBar(DT_STATUSBAR)
'''''''''''''''''''''''''''''''''''AR01''''''''''''''''''''''''''''''''''''
Call SetTcode(DT_SAPTRANSACTIONCODE_OCC1) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call SelectRadioButton("XEINZEL","List assets",False)
Call SelectCheckbox("P_GRID",0,"ON",False)

Call SetTextbox("Company code","BUKRS-LOW","",DT_ABAW_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Asset number","ANLAGE-LOW","",DT_ABAW_0100_ASSET_OCC1,False)
Call SetTextbox("Report date","BERDATUM","",Replace(DT_REPORT_DATE,"/","."),False)
Call SetTextbox("Depreciation area","BEREICH1","",DT_AR01_1000_DEPRECIATION_AREA,False)
Call SetTextbox("Sort Variant","SRTVR","",DT_AR01_1000_SORT_VARIANT,False)
Call TakeScreenShot()
Call PressEnter()    
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call DoubleClickGuiGridCell("",0,1,"Asset",False)

Call SetTextbox("Fiscal year","EDIT_JAHRE","",DT_YEAR,False)
Call PressEnter()   

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")
Call SelectTab("IDC_TABSTRIP","Posted values",False)
Call TakeScreenShot()
Call DoubleClickGuiGridCell("Depreciation posted/planned",0,1,"Depreciation period",False)
Call TakeScreenShot()


Call DoubleClickGuiGridCell("",0,3,"Reference document",False)
wait 3
Call DoubleClickGuiGridCell("Documents in Accounting",0,1,"Object type text",True)

Call TakeScreenShot()

Call VerifyGridCellContent("", 2, "Amount", 0, DT_AMOUNT)



'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()




