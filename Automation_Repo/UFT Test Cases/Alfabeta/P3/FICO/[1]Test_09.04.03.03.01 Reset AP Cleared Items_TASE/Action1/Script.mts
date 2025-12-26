

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.03.03.01 Reset AP Cleared Items
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_09.04.03.03.01 Reset AP Cleared Items"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.03.03.01 Reset AP Cleared Items.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
''
''''''''--------TransactionCode-FB60----------''''
''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_ASSORTMENT_UNIQUE",Cint(DT_ASSORTMENT_UNIQUE)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Clearing Document","RF05R-AUGBL","",DT_FBRA_0100_CLEARING_DOCUMENT,False)
Call SetTextbox("Company Code","RF05R-BUKRS","",DT_FBRA_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05R-GJAHR","",DT_FBRA_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()

Call SelectMenuBar("Clearing;Reset cleared items")
Call ClickButton("Only Reset", True)
Call GetStatusBar("item1","DT_FBRA_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
'Call VerifyStatusBar("Clearing "&DT_FBRA_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" reset")
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_FBRA_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FBRA_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call SetTcode(DT_FBRA_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FBRA_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call SelectRadioButton("X_AISEL","All items",False)
Call ClickButton("Dynamic selections   \(Shift\+F4\)", False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_1,True)


Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_FBRA_0500_CHECK_TOOLTIP_OF_GRIDCELL_0_ICO_AUGP)

Call LogOFF
Call FinalStatus()

