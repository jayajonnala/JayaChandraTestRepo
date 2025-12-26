
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_P2P_01_01_016- Create GR - Create Invoice
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_P2P_01_01_016- Create GR - Create Invoice
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_PRE_P2P_01_01_016- Create GR - Create Invoice"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Input Data\MI\TASE_DT_PRE_P2P_01_01_016- Create GR - Create Invoice.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
'
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot()
'Call CheckTCodeScreen(DT_ME21N_0014_OKCD)


Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call PressEnter() 
Call SetComboByKey("GODYNPRO-REFDOC", DT_MIGO_0010_GODYNPROREFDOC)
'Call SetTextBox("GODEFAULT_TV-BWART","GR goods receipt",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)


Call TakeScreenShot()


Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_MIGO_2000_GODYNPROPO_NUMBER, False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)

Call TakeScreenShot()
Call PressEnter()

Call SetTableData("SAPLMIGOTV_GOITEM","OK",2,"","","ON",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)
'
Call ClickButton("Check Entries   \(F7\)",false)
'Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("btn\[11\]",False)


Call GetStatusBar("item1","DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
'Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT&" posted")


'
Call SetTCode(DT_MIGO_6150_OKCD)     
Call PressEnter()     


Call SetComboByKey("RM08M-VORGANG", DT_MIRO_SELECT_INVOICE)
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIGO_0010_INVOICE_DATE),False)

Call PressEnter()

Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIGO_0010_REFERENCE,False)
Call PressEnter()
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIGO_6020_RM08MREFERENZBELEGTYP)
Call TakeScreenShot()
Call SetTextboxNoLabel("RM08M-EBELN", "", DT_MIGO_6211_RM08MEBELN, False)
Call SetComboByKey("RM08M-XWARE_BNK",DT_MIGO_6211_RM08MXWARE_BNK)

Call PressEnter()
Call SelectCheckbox("INVFO-XMWST", 0, DT_MIGO_0010_CALCULATE_TAX, False)
Call TakeScreenShot()


'Call GetTextboxValue("INVFO-WMWST", "", "DT_MIGO_0010_AMOUNT", false)
Call GetTextboxValue("RM08M-DIFFERENZ", "", "DT_MIGO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT", false)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_MIGO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT),False)
Call PressEnter()
Call TakeScreenShot()


Call ClickButton("Post   \(Ctrl\+S\)",false)


Call GetStatusBar("item1","DT_MIGO_6000_CHECK_TEXT_OF_STATUSBAR_OUTPUT")


'Call VerifyStatusBar("Document no. "&DT_INVOICE_NUMBER& " created")
Call Logoff()
Call FinalStatus()





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




