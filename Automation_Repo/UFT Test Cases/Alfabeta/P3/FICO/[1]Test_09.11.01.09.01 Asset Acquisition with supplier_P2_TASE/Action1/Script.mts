		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.09.02 Asset Acquisition without supplier_P2
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

gstrTestCaseName = "Test_09.11.01.09.01 Asset Acquisition with supplier_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\AB\FICO\DT_09.11.01.09.01 Asset Acquisition with supplier_P2_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-AS01----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FBL1N_1000_VENDOR_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All items",False)
Call TakeScreenShot
Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call SetTextbox("Document Number","%%DYN012-LOW","",DT_FBL1N_1106_DOCUMENT_NUMBER,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)

SAPGuiSession("Session").SAPGuiWindow("Change Layout").SAPGuiToolbar("GridToolbar").PressButton "&FIND"

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_MB51_0841_SEARCH_TERM1,True)
Call TakeScreenShot() 
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot() 

Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)

SAPGuiSession("Session").SAPGuiWindow("Change Layout").SAPGuiToolbar("GridToolbar").PressButton "&FIND" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_MB51_0841_SEARCH_TERM2,True)
Call TakeScreenShot() 
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot() 

Call VerifyGridCellContent("", 1, "Invoice reference", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
'''Call VerifyGridCellContent("", 1,"Document Date", 0, ConvertDate(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1,"Amount in doc. curr.", 0,DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "Local Currency", 0,DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "G/L Account", 0,DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)


Call SetTcode("/n")     
Call PressEnter()
Call SetTcode(DT_FBL1N_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All items",False)
Call TakeScreenShot
Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0,"General Ledger Line Items;Document Number")
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL1N_0100_DOCUMENT_NUMBER,False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Yes",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "Document Number", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", 1, "Document type", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "Posting Key", 0,DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "Amount in local currency", 0,DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Choose Ledger   \(Ctrl\+F3\)",False)
Call TakeScreenShot
Call SetTextbox("Ledger","SVALD-VALUE","",DT_FBL1N_0300_LEDGER,True)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call VerifyGridCellContent("", 1, "Document Number", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", 1, "Document type", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "Posting Key", 0,DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)



Call LogOff
Call FinalStatus ()




