
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_098 -Faire share-substitution_P3_Idoc_PO_Substitution_TASE      
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


gstrTestCaseName = "Test_O2C_08_07_098 _P3_Idoc_PO_Substitution_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_O2C_08_07_098 -Faire share-substitution_P3_Idoc_PO_Substitution.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'
''--------TransactionCode-WE05----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_DATE),False)
Call SetTextbox("to","CREDAT-HIGH","",ConvertDate(DT_DATE),False)
Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE05_1100_MESSAGE_VARIANT,False)
Call SetTextbox("Message Function","MESFCT-LOW","",DT_WE05_1100_MESSAGE_FUNCTION,False)
Call SelectTab("TABSTRIP_IDOCTABBL", "EDI", False)
Call SetTextbox("Transfer File Reference","REFINT-LOW","",DT_WE05_1300_TRANSFER_FILE_REFERENCE,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
If SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=32","attachedtext:=Current Status","name:=EDIDC-STATUS","Index:=0").Exist(5)=False Then 
Call SelectColumnGuiGrid("Selected IDocs", 0,  "Created at", False)
Call ClickButtonToolbar("&SORT_DSC",0)
wait 5
Call SelectRowGuiGridbyRowNo("Selected IDocs", 0,1,False)
'Call DoubleClick()
Call DoubleClickGuiGridCell("Selected IDocs", 0, 1, "IDoc number", False)
End If

Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", "", DT_WE05_0100_CHECK_TEXT_OF_CURRENT_STATUS, False)
Call SelectMenuBar("Goto;Display Links")
Call ActivateCellGuiGridByRefVal("Relationships to IDoc:.*", 0, "Document type", "Purchase Order", "Description", True)
Call SelectRowGuiGrid("Relationships to IDoc:.*", 0, "Document type", "Purchase Order", True)
Call DoubleClickGuiGridCell("Relationships to IDoc:.*", 0, 4, "Document type", True)
Call TakeScreenShot()
''Call GetTextboxValue("MEPO_TOPLINE-EBELN", "", "DT_WE02_1105_CHECK_TEXT_OF_MEPO_TOPLINEEBELN_OUTPUT", False)
Call GetTextboxValue("MEPO_TOPLINE-EBELN", "", "DT_OUTPUT", False)
Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
Call SelectTab("HEADER_DETAIL", "Texts", False)
Call SelectTab("HEADER_DETAIL", "Org. Data", False)
Call TakeScreenShot()
Call VerifyTableCellContent(1, "Article", "SAPLMEGUITC_1211", DT_WE05_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(1, "PO Quantity", "SAPLMEGUITC_1211", DT_WE05_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_0)
Call ClickButton("Messages   \(Shift\+F9\)",False)
Call TakeScreenShot
Call VerifyTableCellContent(2, "Status", "SAPDV70ATC_NAST3", DT_WE05_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_1)
Call VerifyTableCellContent(2, "Output Type", "SAPDV70ATC_NAST3", DT_WE05_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)

''''--------TransactionCode- /nVL10G----------''''

Call SetTcode(DT_WE05_0100_OKCD_OCC1)     
Call PressEnter()     
Call ClickButton("btn\[17\]",False)
Call SetTextbox("Variant","V-LOW","",DT_VARIANT_NAME,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("btn\[8\]",True)
''Call SetTextbox("Deliv. Creation Date","ST_LEDAT-LOW","",ConvertDate(DT_WE05_1000_DELIV_CREATION_DATE),False)
Call SetTextbox("Deliv. Creation Date","ST_LEDAT-LOW","",ConvertDate(DT_WE05_1000_TO),False)
'Call SetTextbox("Shipping Point/Receiving Pt","ST_VSTEL-LOW","",DT_WE05_1000_SHIPPING_POINTRECEIVING_PT,False)
Call SetTextbox("Shipping Point/Receiving Pt","ST_VSTEL-LOW","","",False)
Call SetTextbox("to","ST_LEDAT-HIGH","",ConvertDate(DT_WE05_1000_TO),False)
Call SelectTab("TABSTRIP_ORDER_CRITERIA", "Purchase Orders", False)
Call TakeScreenShot
''Call SetTextbox("Purchasing Document","ST_EBELN-LOW","",DT_WE02_1105_CHECK_TEXT_OF_MEPO_TOPLINEEBELN_OUTPUT,False)
Call SetTextbox("Purchasing Document","ST_EBELN-LOW","",DT_OUTPUT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

''''--------TransactionCode- /nME23N----------''''

Call SetTcode(DT_WE05_0100_OKCD_OCC2)     
Call PressEnter() 
Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call TakeScreenShot
''Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_WE02_1105_CHECK_TEXT_OF_MEPO_TOPLINEEBELN_OUTPUT,True) 
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_OUTPUT,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
Call SelectTab("HEADER_DETAIL", "Texts", False)
Call SelectTab("HEADER_DETAIL", "Org. Data", False)
'Call VerifyTableCellContent(1, "Article", "SAPLMEGUITC_1211", DT_WE05_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_0_OCC1)
'Call VerifyTableCellContent(1, "PO Quantity", "SAPLMEGUITC_1211", DT_WE05_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_0_OCC1)
Call VerifyTableCellContent(1, "Article", "SAPLMEGUITC_1211", DT_WE05_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call VerifyTableCellContent(1, "PO Quantity", "SAPLMEGUITC_1211", DT_WE05_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_1)
'
Call LogOff()

Call FinalStatus ()






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



