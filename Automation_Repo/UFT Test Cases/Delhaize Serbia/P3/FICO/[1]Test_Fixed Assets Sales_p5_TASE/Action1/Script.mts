'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Fixed Assets Sales_p5_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 22th March
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

gstrTestCaseName = "Test_Fixed Assets Sales_p5_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Sales_p6_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
''''----------------------Tcode FB03----------------------------
'
'''''''Enter the Tcode
''''''Call SetTcode(DT_SAPTRANSACTIONCODE) 
''''''Call PressEnter()     ' 
''''''Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
'''''''Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER,False)
''''''Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)
''''''Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB03_0100_FISCAL_YEAR,False)
''''''Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0100_COMPANY_CODE,False)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call PressEnter()
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
'''''''Validate text boxes
''''''Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_FB03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)
''''''Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_FB03_0750_CHECK_TEXT_OF_COMPANY_CODE,False)
'''''''validate grid components
''''''call VerifyGridCellContent("",1,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
''''''call VerifyGridCellContent("",2,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
''''''call VerifyGridCellContent("",1,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
''''''call VerifyGridCellContent("",2,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
''''''Call TakeScreenShot()
''''''
'''''''
'''----------------------Tcode F-92----------------------------
'
'Enter the Tcode
Call SetTcode(DT_FB03_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FB03_0750_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_FB03_0100_DOCUMENT_DATE,"/","."),False)
'Call SetTextbox("Type","BKPF-BLART","",DT_F90_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB03_0100_COMPANY_CODE_OCC1,False)
'Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace(DT_F90_0100_POSTING_DATE,"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB03_0100_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB03_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB03_0100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_FB03_0100_DOCHEADER_TEXT,False)
'Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB03_0100_ACCOUNT,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("BKPF-XMWST",0,DT_FB03_0301_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB03_0301_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB03_0301_TAX_CODE,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_FB03_0301_BUS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB03_0301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB03_0301_ACCOUNT,False)
'Call SetTextbox("TType","RF05A-NEWBW","",DT_F90_0302_TTYPE,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("RF05A-XAABG",0,DT_FB03_0300_ASST_RETIREMENT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB03_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB03_0300_TAX_CODE,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB03_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB03_1007_COST_CENTER,False)

''Call PressEnter()  
'''Capture the screenshot
''Call TakeScreenShot()

Call ClickButton("Display Long Texts for Item",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("ANBZ-XVABG",0,DT_FB03_0210_COMPLETE_RETIREMENT,True)
Call SetTextbox("Asset","RA01B-ANLN1","",DT_FB03_0210_ASSET,True)
Call SetTextbox("Asset Value Date","ANBZ-BZDAT","",Replace(DT_FB03_0210_ASSET_VALUE_DATE,"/","."),True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("1st line","EENO_DYNP-ZEILE","",DT_FB03_1001_1ST_LINE,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy text   \(F5\)",True)'----------------------
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item1","DT_FB03_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB03_0100_CHECK_TEXT_OF_STATUSBAR)

'display saved doc
Call SelectMenuBar("Document;Display")
Wait(1)

'Validate text boxes
Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_FB03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1,False)
Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_FB03_0750_CHECK_TEXT_OF_COMPANY_CODE_OCC1,False)

'validate grid components
call VerifyGridCellContent("",1,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
call VerifyGridCellContent("",2,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
call VerifyGridCellContent("",3,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
call VerifyGridCellContent("",4,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
call VerifyGridCellContent("",5,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
call VerifyGridCellContent("",6,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)

call VerifyGridCellContent("",1,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
call VerifyGridCellContent("",2,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
call VerifyGridCellContent("",3,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
call VerifyGridCellContent("",4,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
call VerifyGridCellContent("",5,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
call VerifyGridCellContent("",6,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)

'Capture the screenshot
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

