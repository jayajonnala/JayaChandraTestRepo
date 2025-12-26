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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Reload DataSheet to updates and calculations
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name       : Test_Interest Calculation_p2_TASE
'.................Author : TCS           :Bitan
'................ Creation Date          : 4th May
'.................Modified By            :
'.................Modified Date/Details  :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Interest Calculation_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Interest Calculation_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'----------------------Tcode FINT----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Wait(5)
Call ClickButtonIfExist("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenShot()

'Call  DoubleClickGuiGridCell("Variant Catalog for Program RFINTITAR","",DT_FINT_0600_GRIDCELL_0_VARIANT_NAME,"Variant name",True)
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RFINTITAR","",DT_FINT_0600_GRIDCELL_0_VARIANT_NAME,True)
Call ClickButtonIfExist("Choose   \(F2\)",True)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FINT_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FINT_1000_COMPANY_CODE,False)
Call TakeScreenShot()

Call SelectCheckbox("A_TESTL","1",DT_FINT_1000_TEST_RUN,False)

Call SetTextbox("Form Date","A_FORMDT","",Replace((DT_FINT_1000_FORM_DATE),"/","."),False)
Call SetTextbox("Interest Calculation To","A_END","",Replace((DT_FINT_1000_INTEREST_CALCULATION_TO),"/","."),False)
Call TakeScreenShot()

Call ClickButtonIfExist("Save as Variant\.\.\.   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("BUTTON_1",True)
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
Wait(4)
Call TakeScreenShot()

SAPGuiSession("name:=ses\[1\]").SAPGuiWindow("program:=SAPLSBAL_DISPLAY").Close
Call TakeScreenShot()
SAPGuiSession("name:=ses\[0\]").SAPGuiWindow("program:=SAPLSLVC_FULLSCREEN","text:=Item Interest Calculation").Close
SAPGuiSession("name:=ses\[0\]").SAPGuiWindow("program:=RFINTITAR","text:=Item Interest Calculation").highlight


Call LogOff()
Call FinalStatus ()

