'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Fixed Assets Disposal_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 17th March
'.................Modified By :
'.................Modified Date/Details :
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrTestCaseName = "Test_Fixed Assets Disposal_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Disposal_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


''----------------------Tcode F-90----------------------------

''''Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_F90_0100_DOCUMENT_DATE,"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F90_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F90_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace(DT_F90_0100_POSTING_DATE,"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F90_0100_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F90_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F90_0100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F90_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F90_0100_ACCOUNT,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call PressEnter()  
Call TakeScreenShot()

Call SelectCheckbox("BKPF-XMWST",0,DT_F90_0302_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0302_AMOUNT,False)
''Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F90_0302_TAX_CODE,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_F90_0302_BUS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F90_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_F90_0302_TTYPE,False)

'Capture the screenshot
'''Call TakeScreenShot()
'''Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F90_0305_TAX_CODE,False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0305_AMOUNT,False)
Call PressEnter()  
'''Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F90_0305_TAX_CODE,False)

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()
'
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item1","DT_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F90_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

'display saved doc
Call SelectMenuBar("Document;Display")
Wait(1)

'getting session details for future
Call GetGridContentByTitle("","","Company Code",1,"DT_F90_0750_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call GetGridContentByTitle("","","Posting Key",1,"DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OUTPUT")
Call GetGridContentByTitle("","","Posting Key",2,"DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OUTPUT")
''''Call GetGridContentByTitle("","","Posting Key",3,"DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL_OUTPUT")
Call GetGridContentByTitle("","","Account",1,"DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OUTPUT")
Call GetGridContentByTitle("","","Account",2,"DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OUTPUT")
'''''Call GetGridContentByTitle("","","Account",3,"DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OUTPUT")
'loading datasheet again
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'validate grid components
call VerifyGridCellContent("",1,"Company Code","",DT_F90_0750_CHECK_TEXT_OF_COMPANY_CODE_OCC1)
call VerifyGridCellContent("",1,"Posting Key","",DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
call VerifyGridCellContent("",2,"Posting Key","",DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
'''''call VerifyGridCellContent("",3,"Posting Key","",DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL_OCC1)
call VerifyGridCellContent("",1,"Account","",DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
call VerifyGridCellContent("",2,"Account","",DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
''''call VerifyGridCellContent("",3,"Account","",DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OCC1)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

