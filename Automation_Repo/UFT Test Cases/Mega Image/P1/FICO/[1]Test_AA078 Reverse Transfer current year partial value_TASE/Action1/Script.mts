
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA078 Reverse Transfer current year partial value
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


gstrTestCaseName = "Test_AA078 Reverse Transfer current year partial value"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA010- Asset Creation - create asset directly in SAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'

''--------TransactionCode-ABo8----------''''
''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Company Code","RLAB01-BUKRS","",DT_AB08_0010_COMPANY_CODE,False)
Call SetTextbox("Asset","RLAB01-ANLN1","",DT_AB08_0010_ASSET,False)
Call SetTextbox("Sub-number","RLAB01-ANLN2","",DT_AB08_0010_SUBNUMBER,False)
Call SetTextbox("Fiscal Year","RLAB01-GJAHR","",Year(Date),False)
Call TakeScreenShot

Call PressEnter()     
Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPLAB01TC_ZAB01_D0100",2,False)
Call TakeScreenShot
Call ClickButton("Reverse Document   \(F6\)",False)
Call TakeScreenShot

Call SetTextbox("Reversal Reason","BAPI6037_REV_DATA-REASON_REV","",DT_AB08_0650_REVERSAL_REASON,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("Line items", 3, "WRBTR", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WRBTR_OCC1)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item2", "DT_AB08_0010_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Asset transaction posted with document no. "& DT_AB08_0010_COMPANY_CODE &" "&DT_AB08_0010_CHECK_TEXT_OF_STATUSBAR_OUTPUT)



''''''--------TransactionCode-AS03----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_AB08_0010_OKCD)     
Call PressEnter()     
Call CheckTCodeScreen(DT_AB08_0010_OKCD)
Call TakeScreenShot
Call PressEnter()   
Call TakeScreenShot

Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("Transactions", 1, "WAERS", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERS_OCC1)

''-------------------------------AS03---------------------------------------------
Call SetTcode(DT_AB08_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AB08_0100_ASSET,False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0,FormatBlank( DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE))


Call LogOff()
Call FinalStatus ()



