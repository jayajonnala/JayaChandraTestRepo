'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.07.02 Change Parked Documents
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

gstrTestCaseName = "Test_09.04.01.07.02 Change Parked Documents"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'----------------------Login----------------------------
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''''''''--------TransactionCode-F-02----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0300_PSTKY,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1007_COST_CENTER,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1007_BUSINESS_AREA,False)
Call TakeScreenShot

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call ClickButton("Park Document   \(Shift\+F4\)",False)
Call GetStatusBar("item1","DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" GR02 was parked")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''''''--------TransactionCode-FBV2----------''''
Call SetTcode(DT_F02_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call PressEnter()     
Call TakeScreenShot

Call FocusTextBoxByIndex("Itm PK  BusA Acct no\.   Description                    Tx", "RF05V-ANZDT", 0, False)
Call TakeScreenShot
Call SendKey("{F2}")
Call TakeScreenShot

Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT_OCC1,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call VerifyStatusBar(Lcase(DT_F02_0100_CHECK_TEXT_OF_STATUSBAR_OCC1))
Call TakeScreenShot

Call PressEnter()     
Call TakeScreenShot

Call FocusTextBoxByIndex("Itm PK  BusA Acct no\.   Description                    Tx", "RF05V-ANZDT", 0, False)
Call TakeScreenShot
Call SendKey("{F2}")
Call TakeScreenShot
Call VerifyTextBoxContent("Text", "BSEG-SGTXT", 0, DT_F02_0300_CHECK_TEXT_OF_TEXT, False)

Call ClickButton("Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call SelectMenuBar("Document;Complete")
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call VerifyStatusBar((DT_F02_0100_CHECK_TEXT_OF_STATUSBAR_OCC2))


'''''''''--------TransactionCode-FBV2----------''''
Call SetTcode(DT_F02_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Post document   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBar((DT_F02_0100_CHECK_TEXT_OF_STATUSBAR_OCC3))

'''''''''--------TransactionCode-FB03----------''''
Call SetTcode(DT_F02_0100_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot

Call PressEnter()     
Call TakeScreenShot

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","G/L Account",True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("", 1, "KOBEZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "Text", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
Call VerifyGridCellContent("", 2, "KOBEZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "MWSKZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 2, "GSBER", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GSBER)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 2, "Functional Area", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_FKBER_LONG)
Call VerifyGridCellContent("", 2, "Text", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)

Call LogOff'
Call FinalStatus()


