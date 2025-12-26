

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.03.03.01 Reset AP Cleared ItemsV2
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

gstrTestCaseName = "Test_09.04.03.03.01 Reset AP Cleared ItemsV2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.03.03.01 Reset AP Cleared ItemsV2.xls"
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
'''''''''--------TransactionCode-FBRA----------''''
'''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_ASSORTMENT_UNIQUE",Cint(DT_ASSORTMENT_UNIQUE)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

' SetTextbox(textboxAttachedText, textboxName, textboxIndex, textboxValue, blnIsItPopup)
Call SetTextbox("Clearing Document","RF05R-AUGBL","",DT_FBRA_0100_CLEARING_DOCUMENT,False)
Call SetTextbox("Company Code","RF05R-BUKRS","",DT_FBRA_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05R-GJAHR","",DT_FBRA_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter
Call TakeScreenShot

Call SelectMenuBar("Clearing;Reset cleared items")
Call ClickButtonIfExist("Reset and Reverse", True)
Call SetTextbox("Reversal Reason","RF05R-STGRD","",DT_FBRA_0300_REVERSAL_REASON,True)
Call SetTextbox("Posting Date","RF05R-BUDAT","",ConvertDate(DT_FBRA_0300_POSTING_DATE),True)
Call SetTextbox("Posting period","RF05R-MONAT","",DT_FBRA_0300_POSTING_PERIOD,True)
Call TakeScreenShot
Call PressEnter
Call TakeScreenShot
Call PressEnter
Call GetTextboxValue("MESSTXT1","","DT_FBRA_0010_CHECK_TEXT_OF_MESSTXT1_OUTPUT",True)
Call PressEnter
Call TakeScreenShot

'''''''''--------TransactionCode-FBL1N----------''''

Call SetTcode(DT_FBRA_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FBRA_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot
Call ClickButton("Dynamic selections   \(Shift\+F4\)", False)
Wait(3)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_2,True)
Call SetTableData("SAPLALDBSINGLE","Single value","4","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_3,True)

Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGPs)

Call LogOff
Call FinalStatus()

