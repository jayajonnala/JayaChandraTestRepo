 
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_042-Returns to EU Vendors_P3   
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
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_P2P_01_01_042-Returns to EU Vendors_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_042-Returns to EU Vendors_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'
'''''''--------TransactionCode-MIRO----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code","RO02")
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTextboxNoLabel("INVFO-BLDAT",0,ConvertDate(DT_MIRO_0010_DOCUMENT_DATE),False)
Call SetComboByKey("RM08M-VORGANG","2")
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call SetTextboxNoLabel("RM08M-EBELN", "", DT_MIRO_6211_RM08MEBELN, False)
Call SetComboByKey("RM08M-XWARE_BNK", DT_MIRO_6211_RM08MXWARE_BNK)


Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_MIRO_0010_POSTING_DATE),False)
Call PressEnter() 

Call SelectCheckbox("INVFO-XMWST", 0, DT_MIRO_0010_CALCULATE_TAX, False)
Call PressEnter() 
Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT",False)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT),False)
Call PressEnter() 
Call TakeScreenShot
Call SelectTab("HEADER", "Details", False)
Call SetComboByKey("Doc. Type", DT_MIRO_0150_DOC_TYPE)
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Simulate Document   \(Ctrl\+Shift\+F7\)",false)
Call TakeScreenShot
Call ClickButtonIfExist("btn\[12\]",true)
Call ClickButton("btn\[11\]",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document no. "&DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")
Call SelectMenuBar("Invoice Document;Display")
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",false)

If VerifyGridCellContentIfExist("Documents in Accounting",1,"Object type text",0,"Accounting document")= True Then
	Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Doc. Number","DT_MIRO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT")
	Call TakeScreenShot()
	Call  ClickButton("Cancel   \(F12\)",True)

Else
	Call GetTextboxValue("BKPF-BELNR",0,"DT_MIRO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
End If


''''''--------TransactionCode-FB03----------''''

Call SetTcode(DT_MIRO_0750_OKCD)     
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_MIRO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_MIRO_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",year(Date),False)
Call PressEnter()
Call TakeScreenShot

Call VerifyTextBoxContent("Document Number", "BKPF-BELNR", 0, DT_MIRO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT, False)
Call VerifyGridCellContent("", 1, "BUKRS", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 4, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call VerifyGridCellContent("", 2, "EBELN", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_EBELN)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

Call LogOff()

Call FinalStatus ()



