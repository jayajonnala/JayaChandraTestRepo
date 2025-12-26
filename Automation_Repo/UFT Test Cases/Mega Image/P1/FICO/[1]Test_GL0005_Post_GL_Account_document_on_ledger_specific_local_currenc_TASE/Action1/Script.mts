
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0005_Post_GL_Account_document_on_ledger_specific_local_currenc_TASE
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


gstrTestCaseName = "Test_GL0005_Post_GL_Account_document_on_ledger_specific_local_currenc_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call Find445CurrentPeriod(DT_TODAY, "DT_PERIOD")


'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'''''''''--------TransactionCode-FB01L----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(Date),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01L_100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01L_100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01L_100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01L_100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(Date),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01L_100_CURRENCYRATE,False)
Call SetTextbox("Ledger Grp","BKPF-LDGRP","",DT_FB01L_100_LEDGER_GRP,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB01L_100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB01L_100_DOCHEADER_TEXT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Text","BSEG-SGTXT","",DT_FB01L_300_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01L_300_AMOUNT,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB01L_1006_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01L_1006_BUSINESS_AREA,False)

Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call SelectCheckbox("BSEG-XNEGP", 0, DT_FB01L_330_NEGATIVE_PSTNG, False)

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01L_330_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01L_330_ACCOUNT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01L_300_AMOUNT_OCC2,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB01L_300_TEXT_OCC2,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01L_330_ACCOUNT,False)

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call SelectMenuBar("Document;Simulate General Ledger")
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FB01L_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FB01L_500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "HKONT", 0, DT_FB01L_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "HKONT", 0, DT_FB01L_500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code RO02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO_OUTPUT",DT_DOC_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)

''''''''''--------TransactionCode-FAGLB03----------''''
'
Call SetTcode(DT_FB01L_100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE","",DT_FB01L_0300_LEDGER,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FB01L_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FB01L_1000_COMPANY_CODE,False)

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("", 0, (Cint(DT_ROW)+1), "Balance", False)

Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)

Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FB01L_1105_DOCUMENT_NUMBER,False)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyifGuiLabelExists(DT_FB01L_120_CHECK_TEXT_OF_L0)
Call VerifyifGuiLabelExists(DT_FB01L_120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_FB01L_120_CHECK_TEXT_OF_NO_NAME_OCC2)
Call VerifyifGuiLabelExists(DT_FB01L_120_CHECK_TEXT_OF_NO_NAME_OCC3)
Call VerifyifGuiLabelExists(DT_FB01L_120_CHECK_TEXT_OF_GL0005SGTXT)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE",0,DT_FB01L_0300_LEDGER_OCC2,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FB01L_1000_ACCOUNT_NUMBER,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Information Message","MESSTXT1", 0, LCase(DT_FB01L_120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA),True)
Call PressEnter()  
Call TakeScreenShot


Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE",0,DT_FB01L_0300_LEDGER_OCC3,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FB01L_1000_ACCOUNT_NUMBER_OCC2,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("", 0, (Month(Date)+1), "Balance", False)

Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)

Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FB01L_1105_DOCUMENT_NUMBER_OCC3,False)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyifGuiLabelExists(DT_FB01L_120_CHECK_TEXT_OF_L0_OCC2)
Call VerifyifGuiLabelExists(DT_FB01L_120_CHECK_TEXT_OF_NO_NAME_OCC4)
Call VerifyifGuiLabelExists(DT_FB01L_120_CHECK_TEXT_OF_NO_NAME_OCC5)
Call VerifyifGuiLabelExists(DT_FB01L_120_CHECK_TEXT_OF_GL0003BKTXT)

Call VerifyifGuiLabelExists_ByIndex(DT_FB01L_120_CHECK_TEXT_OF_NO_NAME_OCC6,0)

Call Logoff'
Call FinalStatus()
