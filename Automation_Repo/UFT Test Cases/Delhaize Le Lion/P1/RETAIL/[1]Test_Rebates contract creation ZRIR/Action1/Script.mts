
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Rebates contract creation ZRIR_TASE
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

gstrTestCaseName = "Test_Rebates contract creation ZRIR_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''' Login '''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''SAP Login'''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)

Call ClickButton("Create Condition Contract   \(F6\)",False)
Call TakeScreenshot
''' SetComboByKey(attachedTextOrComboName, keyValue)
Call SetComboByKey("Condition Contract Type",DT_WCOCO_0100_CONDITION_CONTRACT_TYPE)
Call TakeScreenshot
Call ClickButton("Continue   \(Enter\)",True)
Call SetTextbox("from","KOMWCOCOH-DATE_FROM","",ConvertDate(DT_WCOCO_0510_FROM),False)
Call SetTextbox("to","KOMWCOCOH-DATE_TO","",ConvertDate(DT_WCOCO_0510_TO),False)
Call SetTextbox("Owner Vend","KOMWCOCOH-VEND_OWNER","",DT_WCOCO_0710_OWNER_VEND,False)
Call TakeScreenshot
Call PressEnter()
Call SetTextbox("External Number","KOMWCOCOH-EXT_NUM","",DT_WCOCO_0621_EXTERNAL_NUMBER,False)
Call PressEnter()
Call TakeScreenshot
Call SelectTab("TABSTRIP","Purch.",False)
Call TakeScreenshot
Call SetTextbox("Purch\. Organization","KOMWCOCOH-EKORG","",DT_WCOCO_0623_PURCH_ORGANIZATION,False)
Call SetTextbox("Purchasing Group","KOMWCOCOH-EKGRP","",DT_WCOCO_0623_PURCHASING_GROUP,False)
Call SetTextbox("Company Code","KOMWCOCOH-BUKRS","",DT_WCOCO_0623_COMPANY_CODE,False)
Call TakeScreenshot
Call PressEnter()
Call SelectTab("TABSTRIP","Business volume base",False)
Call TakeScreenshot


'SAPGuiSession("Session").SAPGuiWindow("Create DLL: Reb. based").InsightObject("InsightObject").Click
Call CLickButtonToolBar( "WB2R_NEW_LINE",1)

''' ClickLabel(labelContent, labelIndex, blnIsItPopup)
Call ClickLabel("CV05","0",True)
Call TakeScreenshot
Call ClickButtonIfExist("Copy   \(Enter\)",True)
Call TakeScreenshot
'''' SetGridData(gridTitle, gridRowNumber, gridColumnName, cellValue, blnIsItPopup)
Call SetGridData("Business Volume Base","1","Supplier",DT_WCOCO_0628_GRIDCELL_0_VENDOR,false)
Call SetGridData("Business Volume Base","1","Purchasing Group",DT_WCOCO_0628_GRIDCELL_0_PGR,false)
Call SetGridData("Business Volume Base","1","Company Code",DT_WCOCO_0628_GRIDCELL_0_COCD,false)
Call TakeScreenshot
Call PressEnter()
Call SelectTab("TABSTRIP","Settlement data",False)
Call PressEnter()
Call TakeScreenshot
Call SelectTab("TABSTRIP","Settlement Calendar",False)
Call TakeScreenshot
'SAPGuiSession("Session").SAPGuiWindow("Create DLL: Reb. based").InsightObject("InsightObject_2").Click
Call ClickButtonToolBar("WB2R_CAL_GENERATE", 1)

Call SetTextbox("Valid from","SVALD-VALUE","",ConvertDate(DT_WCOCO_0300_VALID_FROM),True)
Call SetTextbox("Valid to","SVALD-VALUE","",ConvertDate(DT_WCOCO_0300_VALID_TO),True)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot


 Call ClickButtonToolBar("CO_ADD", 0)

''SetGridData(gridTitle, gridRowNumber, gridColumnName, cellValue, blnIsItPopup)
Call TakeScreenshot
Call SetGridData("Conditions","1","Supplier",DT_WCOCO_0510_GRIDCELL_0_VENDOR,false)
Call SetGridData("Conditions","1","Article",DT_WCOCO_0510_GRIDCELL_0_ARTICLE,false)
Call SetGridData("Conditions","1","Purch. Organization",DT_WCOCO_0510_GRIDCELL_0_PORG,false)
Call SetGridData("Conditions","1","Condition Type",DT_WCOCO_0510_GRIDCELL_0_CNTY,false)
Call TakeScreenshot
Call SetGridData("Conditions","1","Condition Rate",DT_WCOCO_0510_GRIDCELL_0_CONDITION_RATE,false)
Call SetGridData("Conditions","1","Valid From",ConvertDate(DT_WCOCO_0510_GRIDCELL_0_VALID_FROM),false)
Call SetGridData("Conditions","1","Valid to",ConvertDate(DT_WCOCO_0510_GRIDCELL_0_VALID_TO),false)
Call PressEnter()
Call TakeScreenshot
Call ClickButton("Release Document Header   \(Shift\+F1\)",False)
Call TakeScreenshot
Call VerifyTextBoxContentIconName("A_RELEASE_ICON","0",DT_WCOCO_0510_CHECK_TOOLTIP_OF_A_RELEASE_ICON,False)
' VerifyTextBoxContentIconName(textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenshot
Call GetGridContentByRefColumn("Messages","0","Message Type","S","Line Content","DT_WCOCO_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TEXT_OUTPUT")
Call VerifyGridCellContent("Messages","1","Line Content","0",DT_WCOCO_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TEXT_OUTPUT)
Call TakeScreenshot
Call ClickButton("Continue   \(Enter\)",True)
Call Takescreenshot()

Call LogOff()
Call FinalStatus ()

