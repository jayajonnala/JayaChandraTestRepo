'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Upload AP Accounting Document_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 23th Feb
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Upload AP Accounting Document_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Upload AP  Accounting Document_p2.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Create Document List/Find Documents   \(Shift\+F8\)",False)
'Capture the screenshot
Call TakeScreenShot()
Dim OnOffStatus_DT
OnOffStatus_DT = "ON"
If DT_FB03_1000_OWN_DOCUMENTS_ONLY = "false" Then
	OnOffStatus_DT = "OFF"
End If
Call SelectCheckbox("UNAME",0,OnOffStatus_DT,False)
'SetTextbox(textboxAttachedText,textboxName,textboxIndex,textboxValue,blnIsItPopup)
Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_FB03_1000_COMPANY_CODE,False)
Call SetTextbox("Entry date","BR_CPUDT-LOW","",Replace((DT_FB03_1000_ENTRY_DATE),"/","."),False)
Call SetTextbox("Ledger","BR_RLDNR-LOW","",DT_FB03_1000_LEDGER,False)
Call SetTextbox("Document type","BR_BLART-LOW","",DT_FB03_1000_DOCUMENT_TYPE,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()
call SelectColumnGuiGrid("","","Time of Entry",False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Sort in Ascending Order   \(Ctrl\+F4\)",False)
'Capture the screenshot
Call TakeScreenShot()
'---------------------validating details for first Document-----------------------------------
call SelectCellGuiGrid("","",DT_FB03_0500_GRIDCELL_0_DOCUMENTNO,"Document Number",False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Choose   \(F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

call VerifyGridCellContent("",1,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
call VerifyGridCellContent("",3,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
call VerifyGridCellContent("",4,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
call VerifyGridCellContent("",1,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
call VerifyGridCellContent("",3,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
call VerifyGridCellContent("",4,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
call VerifyGridCellContent("",1,"Description","",Lcase(DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ))
call VerifyGridCellContent("",2,"Description","",Lcase(DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ))
call VerifyGridCellContent("",3,"Description","",Lcase(DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KOBEZ))
call VerifyGridCellContent("",4,"Description","",Lcase(DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KOBEZ))
call VerifyGridCellContent("",1,"Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
call VerifyGridCellContent("",2,"Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
call VerifyGridCellContent("",3,"Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
call VerifyGridCellContent("",4,"Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AZBET)
call VerifyGridCellContent("",2,"Profit Center","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
call VerifyGridCellContent("",3,"Profit Center","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
'---------------------validating details for second Document-----------------------------------
call SelectCellGuiGrid("","",DT_FB03_0500_GRIDCELL_1_DOCUMENTNO,"Document Number",False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Choose   \(F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

call VerifyGridCellContent("",1,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
call VerifyGridCellContent("",2,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
call VerifyGridCellContent("",3,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL_OCC1)
call VerifyGridCellContent("",1,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
call VerifyGridCellContent("",2,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
call VerifyGridCellContent("",3,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OCC1)
call VerifyGridCellContent("",1,"Description","",Lcase(DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ_OCC1))
call VerifyGridCellContent("",2,"Description","",Lcase(DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ_OCC1))
call VerifyGridCellContent("",3,"Description","",Lcase(DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KOBEZ_OCC1))
call VerifyGridCellContent("",1,"Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
call VerifyGridCellContent("",2,"Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)
call VerifyGridCellContent("",3,"Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET_OCC1)
call VerifyGridCellContent("",2,"Profit Center","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


