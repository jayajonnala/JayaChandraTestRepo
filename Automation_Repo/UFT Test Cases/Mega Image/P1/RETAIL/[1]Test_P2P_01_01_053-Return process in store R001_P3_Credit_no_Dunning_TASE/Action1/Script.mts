
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_060-Copy Purchasing Info Records - transfer inactive article to new vendor
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_P2P_01_01_0279-Duplicate invoice check_P1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_P2P_01_01_053-Return process in store R001_P3_Credit_no_Dunning_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\TASE_DT_P2P_01_01_0279-Duplicate invoice check_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()     

Call TakeScreenShot()
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_DOCUMENT_DATE),False)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call TakeScreenShot
Call SetCombo("RM08M-VORGANG", "Credit Memo")
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call PressEnter() 
Call TakeScreenShot
Call SetTextBoxNoLabel("RM08M-LFSNR",0,DT_MIRO_6212_RM08MLFSNR,False)
Call PressEnter()
wait 5
Call SelectRowGuiTableByRow("SAPLMR1MTC_MSEL_VEN", 1, True)
Call ClickButton("Continue   \(F8\)",True)

Call SelectCheckBox("INVFO-XMWST","0","ON",False)

Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT,False)
Call PressEnter() 
Call TakeScreenShot
Call SelectTab("HEADER","Details",False)
Call SetComboByKey("INVFO-BLART", DT_MIRO_0150_DOC_TYPE)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",false)
Call GetStatusBar("text","DT_MIRO_6000_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar(DT_MIRO_6000_CHECK_TEXT_OF_STATUSBAR_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIRO_6000_OKCD)     
Call PressEnter()
Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_CREDIT_MEMO_NUMBER,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIRO_6150_FISCAL_YEAR,False)
Call PressEnter()
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",false)

Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call GetTextboxValue("BKPF-BELNR",0,"DT_MIRO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)

Call SetTcode(DT_MIRO_0750_OKCD)     
Call PressEnter()

Call SetTextboxNoLabel("LIFNR-LOW","",DT_MIRO_1000_VENDOR,False)
Call SetTextbox("Purch\. Organization","EKORG-LOW","",DT_MIRO_1000_PURCH_ORGANIZATION,False)
Call SetTextbox("Purchasing Document","EBELN-LOW","",DT_MIRO_1000_PURCHASING_DOCUMENT,False)
Call SetTextbox("Partner Function","S_PARVW-LOW","",DT_MIRO_1000_PARTNER_FUNCTION,False)
Call SetTextbox("Movement type","S_BWART-LOW","",DT_MIRO_1000_MOVEMENT_TYPE,False)
Call SetTextbox("Company Code","S_BUKRS-LOW","",DT_MIRO_1000_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Execute   \(F8\)",false)
Call TakeScreenShot
''Call checkifGuilabelExists(DT_MIRO_0120_CHECK_TEXT_OF_LIST_DOES_NOT_CONTAIN_ANY_DATA)
Call LogOff()

Call FinalStatus()




'''''
'''''Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
'''''
'''''Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
'''''
'''''Call TakeScreenShot()
'''''
'''''Call SetTextBox("GR goods receipt","GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
'''''
'''''Call TakeScreenShot()
'''''
'''''Call SelectTab("TS_GOITEM","Article",False)
'''''
'''''Call SetTextBox("Article","GOITEM-MAKTX",0,DT_MIGO_0310_MOVEMENT_TYPE,False)
'''''
'''''Call SelectTab("TS_GOITEM","Quantity",False)
'''''
'''''Call SetTextBox("Qty in Unit of Entry","GOITEM-ERFMG",0,DT_MIGO_0315_MOVEMENT_TYPE,False)
'''''
'''''Call PressEnter()
'''''
'''''Call SelectTab("TS_GOITEM","Where",False)
'''''
'''''Call SetTextBox("Movement type","GOITEM-BWART",0,DT_MIGO_0325_MOVEMENT_TYPE,False)
'''''
'''''Call SetTextBox("Site","GOITEM-NAME1",0,DT_MIGO_0325_SITE,False)
'''''
'''''Call SetTextBox("Storage Location","GOITEM-LGOBE",0,DT_MIGO_0325_STORAGE_LOCATION,False)
'''''
'''''Call PressEnter()
'''''
'''''Call TakeScreenShot()
'''''
'''''Call SelectTab("TS_GOITEM","Partner",False)
'''''
'''''Call SetTextBox("Vendor","GOITEM-VENDORNAME",0,DT_MIGO_0340_VENDOR,False)
'''''
'''''Call PressEnter()
'''''
'''''Call TakeScreenShot()
'''''
'''''Call ClickButton("Check Entries   \(F7\)",False)
'''''
'''''Call TakeScreenShot()
'''''
'''''
'''''Call TakeScreenShot()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''
'''''
'''''
'''''
'''''Call ClickButton("OK_NEXT_ITEM",False)
'''''
'''''Call SelectTab("TS_GOITEM","Article",False)
'''''
'''''Call SetTextBox("Article","GOITEM-MAKTX",0,DT_MIGO_0310_MOVEMENT_TYPE_OCC1,False)
'''''
'''''Call SelectTab("TS_GOITEM","Quantity",False)
'''''
'''''Call SetTextBox("Qty in Unit of Entry","GOITEM-ERFMG",0,DT_MIGO_0315_MOVEMENT_TYPE_OCC1,False)
'''''
'''''Call PressEnter()
'''''
'''''Call SelectTab("TS_GOITEM","Where",False)
'''''
'''''Call SetTextBox("Movement type","GOITEM-BWART",0,DT_MIGO_0325_MOVEMENT_TYPE_OCC1,False)
'''''
'''''Call SetTextBox("Site","GOITEM-NAME1",0,DT_MIGO_0325_SITE_OCC1,False)
'''''
'''''Call SetTextBox("Storage Location","GOITEM-LGOBE",0,DT_MIGO_0325_STORAGE_LOCATION_OCC1,False)
'''''
'''''Call PressEnter()
'''''
'''''Call TakeScreenShot()
'''''
'''''Call SelectTab("TS_GOITEM","Partner",False)
'''''
'''''Call SetTextBox("Vendor","GOITEM-VENDORNAME",0,DT_MIGO_0340_VENDOR_OCC1,False)
'''''
'''''Call PressEnter()
'''''
'''''Call TakeScreenShot()
'''''
'''''Call ClickButton("Check Entries   \(F7\)",False)
'''''
'''''
'''''Call TakeScreenShot()
'''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''
'''''Call ClickButton("Post Document   \(Shift\+F11\)",False)
'''''
'''''Call TakeScreenShot()
'''''
'''''Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
'''''
'''''Call VerifyStatusBar("Article Document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" posted")
'''''
'''''Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION_OCC1)
'''''
'''''Call TakeScreenShot()
'''''
'''''Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)
'''''
'''''Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_2010_GODYNPRODOC_YEAR,False)
'''''
'''''Call PressEnter()
'''''
'''''Call TakeScreenShot()
'''''
'''''Call SelectTab("TS_GOITEM","Purchase Order Data",False)
'''''
'''''Call GetTextboxValue("GOITEM-EBELN", 0, "DT_MIGO_0320_CHECK_TEXT_OF_PURCHASE_ORDER_OUTPUT", False)
'''''
'''''Call TakeScreenShot()
'''''








'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




