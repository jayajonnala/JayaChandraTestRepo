		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.02.01.03 Clear AR Accounts (Manual and Automatic)


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

gstrTestCaseName = "Test_09.06.02.01.03 Clear AR Accounts (Manual and Automatic)"
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

'''''--------TransactionCode-F.21----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BUKRX-LOW","",DT_FB70_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","GJAHX-LOW","",DT_FB70_1000_FISCAL_YEAR,False)
Call TakeScreenShot

Call ClickButton("%_DOCNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB70_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FB70_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot

Call ClickButtonIfExist("Copy   \(F8\)",True)
Call TakeScreenShot
Call SelectCheckbox("X_KUNNR", 0, "ON", False)
Call TakeScreenShot
Call SetTextbox("Posting Date","POSTDATE-LOW","",ConvertDate(DT_FB70_1000_POSTING_DATE),False)
Call TakeScreenShot
Call SetTextbox("Customers","KONTD-LOW","",DT_FB70_1000_CUSTOMERS,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)",False)
Call SelectCheckbox("X_TESTL", 0, "OFF", False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call GetStatusBar("item1","DT_FB70_0120_CHECK_DOC_NUMBER_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB70_0120_CHECK_DOC_NUMBER_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB70_0120_CHECK_DOC_NUMBER_OUTPUT",DT_DOC)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyifGuiLabelExistsByRelativeid(DT_FB70_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[1,10\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FB70_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[1,11\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FB70_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"wnd\[0\]/usr/lbl\[27,12\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FB70_0120_CHECK_TEXT_OF_NO_NAME_OCC3,"wnd\[0\]/usr/lbl\[47,12\]")

Call Logoff'
Call FinalStatus()



