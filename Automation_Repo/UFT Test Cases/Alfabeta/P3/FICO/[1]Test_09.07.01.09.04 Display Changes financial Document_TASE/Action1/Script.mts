		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.09.04 Display Changes financial Document
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

gstrTestCaseName = "Test_09.07.01.09.04 Display Changes financial Document"
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

'''''--------TransactionCode-FB04 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     

Call SetTextbox("Document Number","RF01A-BELNR","",DT_FB04_0107_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF01A-BUKRS","",DT_FB04_0107_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF01A-GJAHR","", YEar(DT_FB04_0107_FISCAL_YEAR), False)
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("All changes   \(F6\)",False)

Call VerifyifGuiLabelExists(Lcase(DT_FB04_0120_CHECK_TEXT_OF_CORRECTION_FOR_RENTING))
Call VerifyifGuiLabelExists(Lcase(DT_FB04_0120_CHECK_TEXT_OF_CHANGED_ON))
Call VerifyifGuiLabelExists(ConvertDate(DT_FB04_0120_CHECK_TEXT_OF_NO_NAME))
Call VerifyifGuiLabelExists(Lcase(DT_FB04_0120_CHECK_TEXT_OF_CORRECTION_OF_RENT))

Call LogOff()
Call FinalStatus()


