		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.02.04.01 Reset Cleared AR Items

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

gstrTestCaseName = "Test_09.06.02.04.01 Reset Cleared AR Items"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''''--------TransactionCode-FBRA----------''''
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
Call ClickButton("Reset cleared items   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButton("Only Reset",True)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

'''''''--------TransactionCode-FBL5N----------''''

Call SetTcode(DT_FBRA_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBRA_1000_CUSTOMER_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All items", False)
Call SetTextbox("Posting date","SO_BUDAT-LOW","",ConvertDate(DT_FBRA_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_FBRA_1000_TO),False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)
Call TakeScreenShot

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_2,True)
Call TakeScreenShot

Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 3, "BELNR", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BELNR)
Call VerifyGridCellContent("", 4, "BELNR", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BELNR)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BLART", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 3, "BLART", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
Call VerifyGridCellContent("", 4, "BLART", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLART)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 2, "ICO_AUGP", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ICO_AUGP)
Call VerifyGridCellContent("", 3, "ICO_AUGP", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ICO_AUGP)
Call VerifyGridCellContent("", 4, "ICO_AUGP", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ICO_AUGP)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call VerifyGridCellContent("", 4, "DMSHB", 0, DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMSHB)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 2, "BLDAT", 0, ConvertDate(DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
Call VerifyGridCellContent("", 3, "BLDAT", 0, ConvertDate(DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLDAT))
Call VerifyGridCellContent("", 4, "BLDAT", 0, ConvertDate(DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLDAT))
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()




