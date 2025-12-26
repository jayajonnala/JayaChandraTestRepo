	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrTestCaseName = "Test_PRE_4_09.04.03.02.03 PostWithClearing_GlIndicator X (510100-xxx)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_PRE_4_09.04.03.02.03 PostWithClearing_GlIndicator X (510100-xxx)_Output.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------MIGO -----------------------------------

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_MIRO_1000_COMPANY_CODE,False)     ' - Line (19)
call ClickButton("Continue   \(Enter\)",fALSE)

call SetCombo("RM08M-VORGANG","Invoice")
call ClickButton("btn\[12\]",fALSE)

''invoice date - 
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace(DT_MIRO_0010_INVOICE_DATE,"/","."),False)     ' - Line (19)
''Reference - 
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)     ' - Line (19)
'call SetComboByKey("M08M-REFERENZBELEGTYP",DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace(DT_MIRO_0010_POSTING_DATE,"/","."),False)     ' - Line (19)
'''purcasing doc -
Call SetTextboxNoLabel("RM08M-EBELN","",DT_MIRO_6211_RM08MEBELN,False)
Call TakeScreenShot
'call SetComboByKey("RM08M-XWARE_BNK",DT_MIRO_6211_RM08MXWARE_BNK)
'call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)
'TakeScreenShot()
'''verify balance amt
'call VerifyTextBoxContent("Balance","RM08M-DIFFERENZ","",replace(DT_MIRO_6000_CHECK_TEXT_OF_BALANCE,"$",""),False)
Call PressEnter()
Call TakeScreenShot
' GetTextboxValue(textboxName, textboxIndex, dataTableColumnName, blnIsItPopup)
call GetTextboxValue("RM08M-DIFFERENZ",1,"DT_Amount_output",False)
Call TakeScreenShot
Call SetTextbox("Amount","INVFO-WRBTR","",Replace(DT_Amount_output,"-",""),False)     ' - Line (19)
call SetTableData("SAPLMR1MTC_MR1M","Amount","1","","",replace(DT_Amount_output,"-",""),false)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
'call VerifyTableCellContent(1,"Amount","SAPLMR1MTC_MR1M",(replace(DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_AMOUNT_0,"-","")))
call VerifyTableCellContent(1,"Quantity","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_QUANTITY_0)
call VerifyTableCellContent(1,"Order Unit","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_ORDER_UNIT_0)
call VerifyTableCellContent(1,"Purchase Order","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_0)
'DT_MIRO_6310_CHECK_VALUE_OF_TABLECELL_TAX_CODE_0
'call VerifyTableCellContent(1,"Tax Code","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_VALUE_OF_TABLECELL_TAX_CODE_0)
'D0 \(GR VAT 0,0% Intracom Purchases - Trade Goods\)
Call TakeScreenShot
Call ClickButton("btn\[11\]",False)
Call TakeScreenShot
'''Ouput'''

Call GetStatusBar("item2","DT_PO_NUMBER_Output")
call VerifyStatusBar("created")
Call TakeScreenShot
'''MIRO CODE'''''

'''Enter input fields'''
'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_MIRO_1000_COMPANY_CODE,False)     ' - Line (19)
'call ClickButton("Continue   \(Enter\)",fALSE)
'''invoice date - 
'Call SetTextbox("Invoice date","INVFO-BLDAT","",DT_MIRO_0010_INVOICE_DATE,False)     ' - Line (19)
'''Reference - 
'Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)     ' - Line (19)
'''posting date - 
'Call SetTextbox("Posting Date","INVFO-BUDAT","",DT_MIRO_0010_POSTING_DATE,False)     ' - Line (19)
''''purcasing doc -
'call SetTextboxByGuiLabelXYCordDiff("Company Code","",-7,132,DT_MIRO_6211_RM08MEBELN) 
'Call PressEnter()
'''Amount
'
'call SetTableData("SAPLMR1MTC_MR1M","Amount","1","","",DT_MIRO_0010_AMOUNT,false)
'Call PressEnter()
''CALL SendKey("{ENTER}0")
'
'
'' Verifications 
'call VerifyTableCellContent(1,"Amount","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
'call VerifyTableCellContent(1,"Quantity","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_QUANTITY_0)
'call VerifyTableCellContent(1,"Order Unit","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_ORDER_UNIT_0)
'call VerifyTableCellContent(1,"Purchase Order","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_0)
'call VerifyTableCellContent(1,"Tax Code","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_VALUE_OF_TABLECELL_TAX_CODE_0)
'Call ClickButton("Save   (Ctrl+S)",False)
'
''''Ouput'''
'
'If VerifyStatusBar("Created")= True Then
'	Call GetStatusBar("item1","DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
'End If 
'
'
'''Closing application'''
Call LogOff() 
FinalStatus()







'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


