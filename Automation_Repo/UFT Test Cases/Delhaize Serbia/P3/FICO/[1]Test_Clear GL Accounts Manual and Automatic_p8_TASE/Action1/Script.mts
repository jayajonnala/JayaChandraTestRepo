'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Clear GL Accounts Manual and Automatic_p8_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 22th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear GL Accounts Manual and Automatic_p8_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear GL Accounts  Manual and Automatic_p8_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
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

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB03_0100_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Document Header   \(F5\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Document type","BKPF-BLART","",DT_FB03_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)
Call VerifyTextBoxContent("Document type","T003T-LTEXT","",Lcase(DT_FB03_1710_CHECK_TEXT_OF_DOCUMENT_TYPE_OCC1),True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue/Confirm   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

'verify the grid coponents
call VerifyGridCellContent("",1,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

call VerifyGridCellContent("",1,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

call VerifyGridCellContent("",1,"Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
call VerifyGridCellContent("",2,"Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_FB03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

