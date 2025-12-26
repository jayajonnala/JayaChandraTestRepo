		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.02.02.02 Reset GL Cleared Items (with reversal)
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

gstrTestCaseName = "Test_09.07.02.02.02 Reset GL Cleared Items (with reversal)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'''''''--------------FBRA----------------'''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_ASSORTMENT_UNIQUE",Cint(DT_ASSORTMENT_UNIQUE)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Clearing Document","RF05R-AUGBL","",DT_FBRA_0100_CLEARING_DOCUMENT,False)
Call SetTextbox("Company Code","RF05R-BUKRS","",DT_FBRA_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05R-GJAHR","",Year(DT_FBRA_0100_FISCAL_YEAR),False)
Call TakeScreenShot
Call ClickButton("List Accounts Involved   \(Ctrl\+F1\)",False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("List Cleared Items   \(Ctrl\+F2\)",False)
Call TakeScreenShot
CAll SelectColumnGuiGrid("", 0, "Document Number", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)

Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",True)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FBRA_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",Fasle)
Call ClickButton("Reset cleared items   \(Ctrl\+S\)",Fasle)
Call ClickButton("Reset and Reverse",True)

Call SetTextbox("Reversal Reason","RF05R-STGRD","",DT_FBRA_0300_REVERSAL_REASON,True)
Call SetTextbox("Posting Date","RF05R-BUDAT","",ConvertDate(DT_FBRA_0300_POSTING_DATE),True)
Call SetTextbox("Posting period","RF05R-MONAT","",ConvertDoubleDigit(Cstr(Month(DT_FBRA_0300_POSTING_PERIOD))),True)
Call PressEnter()
Call VerifyTextBoxContent("Information Message","MESSTXT1",0,Lcase(DT_FBRA_0010_CHECK_TEXT_OF_MESSTXT1),True)
Call PressEnter()
Call GetTextboxValue("MESSTXT1", 0, "DT_FBRA_0010_CHECK_TEXT_OF_MESSTXT1_OCC1_OUTPUT", True)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FBRA_0010_CHECK_TEXT_OF_MESSTXT1_OCC1_OUTPUT",DT_FBRA_0010_CHECK_TEXT_OF_MESSTXT1_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call PressEnter()

Call LogOff()
Call FinalStatus()


