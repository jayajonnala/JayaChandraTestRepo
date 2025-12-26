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

 '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Actual Postings with Statistical Internal Order_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 12th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Actual Postings with Statistical Internal Order_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Actual Postings with Statistical Internal Order_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'
'''----------------------Tcode F-02----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDateFormat(DT_F02_0100_DOCUMENT_DATE),False)
'Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_F02_0100_DOCUMENT_DATE,"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
'Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace(DT_F02_0100_POSTING_DATE,"/","."),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDateFormat(DT_F02_0100_POSTING_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F02_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1007_COST_CENTER,False)
Call SetTextbox("Order","COBL-AUFNR","",DT_F02_1007_ORDER,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0300_ACCOUNT,False)

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()


Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0302_AMOUNT,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_F02_0302_BUS_AREA,False)
Call SetTextboxNoLabel("BSEG-ZTERM","",DT_F02_0302_PAYT_TERMS,False)
'Call SetTextbox("Bline Date","BSEG-ZFBDT","",Replace(DT_F02_0302_BLINE_DATE,"/","."),False)
Call SetTextbox("Bline Date","BSEG-ZFBDT","",ConvertDateFormat(DT_F02_0302_BLINE_DATE),False)

Call SelectMenuBar("Document;Simulate")
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Company Code","RF05A-AZBUK","",DT_F02_0750_CHECK_TEXT_OF_COMPANY_CODE,False)
Call VerifyTextBoxContent("Currency","BKPF-WAERS","",DT_F02_0750_CHECK_TEXT_OF_CURRENCY,False)

 Call PressEnter() 
Call SelectMenuBar("Document;Simulate General Ledger")

Call ClickButtonIfExist("Change Layout\.\.\.   \(Ctrl\+F8\)",False)
Call SelectRowGuiGridbyRowNo("Column Set","","87",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call SelectRowGuiGridbyRowNo("Column Set","","133",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)


'Call ClickButtonIfExist("Select Layout\.\.\.   \(Ctrl\+F9\)",False)
'Call SetComboByKey("Layout setting","X")
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Posting Key","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",1,"G/L Account","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",1,"Amount","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",1,"Tax Code","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("",1,"Order","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUFNR)

Call VerifyGridCellContent("",2,"Posting Key","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",2,"Amount","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",2,"Tax Code","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)

Call ClickButtonIfExist("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Goto;Last Item")
'Capture the screenshot
Call TakeScreenShot()
''Call VerifyTextBoxContent("Vendor","LFA1-LIFNR","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR,False)

Call VerifyTextBoxNoLabelContent("LFA1-LIFNR","",DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR,False)

Call ClickButtonIfExist("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

