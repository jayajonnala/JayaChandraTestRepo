

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_1_01.02.01.05.02 Post_Reverse Actual Cost Reposting_TASE
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

gstrTestCaseName = "Test_PRE_1_01.02.01.05.02 Post_Reverse Actual Cost Reposting_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


'''''----------------------Login----------------------------'''''
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'----------------------Tcode FAGLB03----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call TakeScreenShot()
Call PressEnter() 
Wait(2)
Call TakeScreenShot()

'''''''''
Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)
Call ClickButton("Execute   \(F8\)",fALSE)
Wait(2)
Call TakeScreenShot()
'

Call GetGridContentByTitle("",0,"Debit",DT_INDEX,"DT_FAGLB03_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DEBIT_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FAGLB03_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DEBIT_OUTPUT",(DT_FAGLB03_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DEBIT))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetGridContentByTitle("",0,"Credit",DT_INDEX,"DT_FAGLB03_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_CREDIT_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FAGLB03_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_CREDIT_OUTPUT",(DT_FAGLB03_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_CREDIT))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetGridContentByTitle("",0,"Balance",DT_INDEX,"DT_FAGLB03_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BALANCE_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FAGLB03_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BALANCE_OUTPUT",(DT_FAGLB03_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BALANCE))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()




