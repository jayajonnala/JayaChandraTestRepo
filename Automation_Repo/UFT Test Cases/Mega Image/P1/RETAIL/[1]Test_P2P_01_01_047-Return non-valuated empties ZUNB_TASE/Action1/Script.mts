
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_P2P_01_01_047-Return non-valuated empties ZUNB
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




gstrTestCaseName = "Test_P2P_01_01_047-Return non-valuated empties ZUNB"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\TASE_DT_P2P_01_01_047-Return non-valuated empties ZUNB.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
''''''''''MB01
Call SetTcode(DT_SAPTRANSACTIONCODE)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Storage Location","RM07M-LGORT","",DT_MB01_0200_STORAGE_LOCATION,False)

Call SetTextBox("Site","RM07M-WERKS","",DT_MB01_0200_SITE,False)

Call SetTextBox("Movement Type","RM07M-BWARTWE","",DT_MB01_0200_MOVEMENT_TYPE,False)

Call SetTextBox("Posting Date","MKPF-BUDAT","",ConvertDate(DT_MB01_0200_POSTING_DATE),False)

Call SetTextBox("Document Date","MKPF-BLDAT","",ConvertDate(DT_MB01_0200_DOCUMENT_DATE),False)

Call TakeScreenShot()

Call PressEnter()

''Call SetTextBox("Vendor","MSEGK-LIFNR",0,DT_MB01_0231_VENDOR,False)
Call SetTextBoxNoLabel("MSEGK-LIFNR",0,DT_MB01_0231_VENDOR,False)

Call PressEnter()

Call SetTextBox("Article","MSEG-MATNR",0,DT_MB01_0231_ARTICLE,False)

'Call PressEnter()

Call SetTextBox("Quantity","MSEG-ERFMG",0,DT_MB01_0231_QUANTITY,False)


Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_MB01_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT")

Call VerifyStatusBar("Document "&DT_MB01_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT&" posted")

Call TakeScreenShot()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call SetTcode(DT_MB01_0200_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MB01_0010_GODYNPROACTION)

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MB01_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT, False)

Call TakeScreenShot()



Call PressEnter()

Call SelectTab("TS_GOHEAD","General",False)

Call GetTextboxValue("GOHEAD-LFSNR", 0, "DT_MB01_0110_CHECK_TEXT_OF_DELIVERY_NOTE_OUTPUT", False)

Call VerifyTableCellContent(1, "Movement type", "SAPLMIGOTV_GOITEM", DT_MB01_0325_CHECK_TEXT_OF_MOVEMENT_TYPE)

Call VerifyTableCellContent(1, "Stor. Location", "SAPLMIGOTV_GOITEM", DT_MB01_0325_CHECK_TEXT_OF_STORAGE_LOCATION)

Call VerifyTableCellContent(1, "Article", "SAPLMIGOTV_GOITEM", DT_MB01_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)

Call SelectTab("TS_GOITEM","Article",False)

Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Quantity",False)

Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Where",False)

Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Partner",False)

Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Output",False)

Call TakeScreenShot()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_MB01_0001_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Article Document","RG_MBLNR-LOW",0,DT_MB01_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT,False)

Call SetTextBox("Article Doc\. Year","PM_MJAHR",0,DT_MB01_1000_ARTICLE_DOC_YEAR,False)

Call SetTextBox("Processing Mode","PM_VERMO",0,DT_MB01_1000_PROCESSING_MODE,False)

Call TakeScreenShot()

Call PressEnter()

Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()

Call  SelectCheckboxNoLabel(1, DT_MB01_0120_NO_NAME, False)

Call TakeScreenShot()

Call ClickButton("Print preview   \(Shift\+F4\)",False)

Call TakeScreenShot()

Call LogOff()

Call FinalStatus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





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




