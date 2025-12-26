

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.04.04 Adjustment Postings Without Clearing
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

gstrTestCaseName = "Test_09.04.01.04.04 Adjustment Postings Without Clearing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.04.04 Adjustment Postings Without Clearing.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''


''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'''--------TransactionCode-F-42 ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB01_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call SetTextbox("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_FB01_0100_DOCUMENT_DATE), False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB01_0100_REFERENCE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0302_PSTKY,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB01_0302_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0302_AMOUNT,False)
Call SetTextbox("G/L Acc","BSEG-HKONT","",DT_FB01_0302_GL_ACC,False)
Call TakeScreenShot
Call TakeScreenShot
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0302_ACCOUNT,False)
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB01_0300_TAX_CODE,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB01_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1007_BUSINESS_AREA,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB01_0300_TEXT,False)
Call TakeScreenShot
Call PressEnter()

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02" )

Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot

'Call SelectMenuBar("Document;Display")
'Call TakeScreenShot
'
''''''--------TransactionCode-FAGLL03 ----------''''
Call SetTcode(DT_FB01_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FB01_1000_GL_ACCOUNT,False)
Call TakeScreenShot
Call ClickButtonIfExist("Custom Selections   \(Ctrl\+F1\)",False)
Call TakeScreenShot

Call ActivateNodeGuiTree(4,"General Ledger Line Items;Document Number")
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FB01_0100_DOCUMENT_NUMBER,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot


Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "Local Currency", 0, DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
'Call VerifyGridCellContent("", 1, "HWAER", 0, DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "MWSKZ", 0, DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call LogOff'
Call FinalStatus()
