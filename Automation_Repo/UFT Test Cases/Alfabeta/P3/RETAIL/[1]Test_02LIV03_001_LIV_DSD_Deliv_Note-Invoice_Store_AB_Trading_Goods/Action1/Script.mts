
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_02_053-Eshop comission manual in SAP - correct values
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_02_053-Eshop comission manual in SAP - correct values
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_02LIV03_001_LIV_DSD_Deliv_Note-Invoice_Store_AB_Trading_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_02LIV03_001_LIV_DSD_Deliv_Note-Invoice_Store_AB_Trading_Goods_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'--------------------------ZMDIV_FENTRY---------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Delivery Note/Invoice   \(Shift\+F5\)",false)
Call TakeScreenShot

Call SetTextbox("Posting Date","GS_HEADER-BUDAT","",convertdate(DT_ZMDIV_FENTRY_0201_POSTING_DATE),False)
Call SetTextbox("Document Date","GS_HEADER-BLDAT","",convertdate(DT_ZMDIV_FENTRY_0201_DOCUMENT_DATE),False)
Call SetTextbox("Reference","GS_HEADER-XBLNR","",DT_ZMDIV_FENTRY_0201_REFERENCE,False)
Call SetTextbox("VAT Reg\. No\.","GS_HEADER-STCEG","",DT_ZMDIV_FENTRY_0201_VAT_REG_NO,False)
'Call SetTextbox("Site","GS_HEADER-WERKS","",DT_ZMDIV_FENTRY_0201_SITE,False)
Call SetTextbox("Site",".*WERKS","",DT_ZMDIV_FENTRY_0201_SITE,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Article","GS_NEW_ITEM-IDNLF","",DT_ZMDIV_FENTRY_0202_ARTICLE,False)
Call SetTextbox("Quantity","GS_NEW_ITEM-MENGE","",DT_ZMDIV_FENTRY_0202_QUANTITY,False)
Call TakeScreenShot
'Call SetTextbox("UoM","GS_NEW_ITEM-MEINS","",DT_ZMDIV_FENTRY_0202_UOM,False)
Call PressEnter()
Call FocusTextBox("UoM","GS_NEW_ITEM-MEINS",False)
Call PressEnter()
Call TakeScreenShot

Call GetTableCellData("ZMDIV_FENTRYTC_0400","Purchase Value",1,"","","DT_ZMDIV_FENTRY_0400_GET_TEXT_OF_TABLECELL_PURCHASE_VALUE_0_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTableData("ZMDIV_FENTRYTC_0400","Net Invoice Value",1,"","",DT_ZMDIV_FENTRY_0400_TABLECELL_NET_INVOICE_VALUE_0,False)
Call TakeScreenShot

Call SetTextbox("Article","GS_NEW_ITEM-IDNLF","",DT_ZMDIV_FENTRY_0202_ARTICLE_OCC2,False)
Call PressEnter()
Call SetTextbox("Quantity","GS_NEW_ITEM-MENGE","",DT_ZMDIV_FENTRY_0202_QUANTITY,False)
Call PressEnter()
Call TakeScreenShot

Call PressEnter()
Call FocusTextBox("UoM","GS_NEW_ITEM-MEINS",False)
Call PressEnter()

Call TakeScreenShot
Call GetTableCellData("ZMDIV_FENTRYTC_0400","Purchase Value",2,"","","DT_ZMDIV_FENTRY_0400_GET_TEXT_OF_TABLECELL_PURCHASE_VALUE_1_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTableData("ZMDIV_FENTRYTC_0400","Net Invoice Value",2,"","",DT_ZMDIV_FENTRY_0400_TABLECELL_NET_INVOICE_VALUE_1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call GetTextboxValue("GS_HEADER-WDIFF","","DT_ZMDIV_FENTRY_0201_GET_TEXT_OF_AMOUNT_DIFF_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Gross Amount","GS_HEADER-WRBTR","",DT_ZMDIV_FENTRY_0201_GROSS_AMOUNT,False)
Call TakeScreenShot
Call PressEnter()

Call ClickButton("Pricing conditions",False)
Call TakeScreenShot
Call GetTableCellData("ZMDIV_FENTRYTC_TAX","Tax Amount",1,"","","DT_ZMDIV_FENTRY_0201_TAX_AMOUNT_OUTPUT",True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickButton("Continue   \(F5\)",True)
Call PressEnter()
Call TakeScreenShot

Call VerifyTextBoxContentIconName("GV_TAX_ICON","",DT_ZMDIV_FENTRY_0201_CHECK_ICONNAME_OF_TAX_AMOUNT,False)
Call VerifyTextBoxContentIconName("GV_DIFF_ICON","",DT_ZMDIV_FENTRY_0201_CHECK_ICONNAME_OF_AMOUNT_DIFF,False)
Call TakeScreenShot

Call ClickButton("Post   \(F8\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_ZMDIV_FENTRY_200_CHECK_TEXT_OF_STATUSBAR)

''''Call ClickButton("Bill of material subitem",false)
''''Call ClickButton("Continue   \(Enter\)",true)
''''
''''Call SetTextbox("UoM","GS_NEW_ITEM-MEINS","",DT_ZMDIV_FENTRY_0202_UOM,False)
''''
''''Call PressEnter()
''''Call PressEnter()
''''Call ClickButton("Post   \(F8\)",false)

Call ClickButtonifexist("Continue   \(Enter\)",true)
Call ClickButtonifexist("Yes",true)
Call TakeScreenShot

Call LogOff()
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




