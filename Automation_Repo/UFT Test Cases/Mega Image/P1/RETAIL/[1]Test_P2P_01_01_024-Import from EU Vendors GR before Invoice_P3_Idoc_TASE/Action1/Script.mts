
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_024-Import from EU Vendors GR before Invoice_P3_Idoc
      
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

gstrTestCaseName = "Test_P2P_01_01_024- Invoice_P3_Idoc"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_024-Import from EU Vendors GR before Invoice_P3_Idoc.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''
''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_DATE),False)
Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE05_1100_MESSAGE_VARIANT,False)
Call SetTextbox("Message Function","MESFCT-LOW","",DT_WE05_1100_MESSAGE_FUNCTION,False)
Call PressEnter() 
Call TakeScreenShot
Call SelectTab("TABSTRIP_IDOCTABBL","EDI",False)
Call TakeScreenShot
Call SetTextbox("Transfer File Reference","REFINT-LOW","",DT_WE05_1300_TRANSFER_FILE_REFERENCE,False)
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",false)
Call TakeScreenShot
Call VerifyTextBoxContent("Current Status","EDIDC-STATUS",2,DT_WE05_0100_CHECK_TEXT_OF_CURRENT_STATUS,False)

Call SetTcode(DT_WE05_0100_OKCD)     
Call PressEnter() 
Call SetComboByKey("GODYNPRO-ACTION", DT_WE05_0010_GODYNPROACTION)
Call PressEnter() 
wait 2

Call SetTextBoxNoLabel("GODYNPRO-MAT_DOC","0",DT_WE05_2010_GODYNPROMAT_DOC,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call SelectTab("TS_GOHEAD","General", False)
Call VerifyTextBoxContent("Delivery Note","GOHEAD-LFSNR",1,DT_WE05_0110_CHECK_TEXT_OF_DELIVERY_NOTE,False)
Call ClickButtonIfExist("Close Detail Data",false)
Call VerifyTableCellContent(1,"Article","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(2,"Article","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call VerifyTableCellContent(3,"Article","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_2)
Call VerifyTableCellContent(4,"Article","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_3)
Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)
Call VerifyTableCellContent(1,"Direction","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_0)
Call VerifyTableCellContent(2,"Movement type","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_1)
Call VerifyTableCellContent(2,"Direction","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_1)
Call VerifyTableCellContent(3,"Movement type","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_2)
Call VerifyTableCellContent(3,"Direction","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_2)
Call VerifyTableCellContent(4,"Movement type","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_3)
Call VerifyTableCellContent(4,"Direction","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_3)
Call VerifyTableCellContent(1,"Purchase order","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_0)
Call VerifyTableCellContent(2,"Purchase order","SAPLMIGOTV_GOITEM",DT_WE05_0200_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_1)
Call SelectTab("TS_GOHEAD","Doc. info", False)
Call ClickButtonIfExist("FI Documents",False)
Call TakeScreenShot
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)
Call VerifyGridCellContent("", 1, "BSCHL", 0,DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call ClickButton("Back   \(F3\)",False)
Call GetGridContent("Documents in Accounting",0,"Document Number",2,"Object type text","Accounting document","DT_WE05_0500_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Output",DataRowSet)
Call SelectRowGuiGrid("Documents in Accounting", 0, "Document Number",DT_WE05_0500_OUTPUT, True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
Call ClickButton("Back   \(F3\)",False)

Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Open Detail Data",false)
Call SelectTab("TS_GOITEM","Output", False)
Call TakeScreenShot
Call ClickButton("Display outputs",False)
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_WE05_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(1,"Status","SAPDV70ATC_NAST3",DT_WE05_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call SetTcode(DT_WE05_0100_OKCD_OCC1)     
Call PressEnter() 
Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_WE05_1000_ARTICLE_DOCUMENT,False)
Call SetTextbox("Processing Mode","PM_VERMO","",DT_WE05_1000_PROCESSING_MODE,False)
Call SetTextbox("Article Doc\. Year","PM_MJAHR","",DT_WE05_1000_ARTICLE_DOC_YEAR,False)
Call SetTextbox("Sort order","PM_NSORT","",DT_WE05_1000_SORT_ORDER,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call SelectCheckBoxNoLabel(0,"ON",False)

Call ClickButton("Print preview   \(Shift\+F4\)",False)

Call TakeScreenShot
Call SetTcode(DT_WE05_0120_OKCD)     
Call PressEnter() 
Call SetTextbox("Purchasing Document","EBELN-LOW","",DT_WE05_1000_PURCHASING_DOCUMENT,False)
Call SetTextbox("Purch\. Organization","EKORG-LOW","",DT_WE05_1000_PURCH_ORGANIZATION,False)
Call SetTextboxNoLabel("LIFNR-LOW","",DT_WE05_1000_VENDOR,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call ClickButton("Detail List   \(Ctrl\+Shift\+F7\)",False)
Call ClickButton("Current Display Variant   \(Ctrl\+F8\)",False)
Call SelectCellGuiGrid("Column Set",0,0,"Column Name",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SelectCellGuiGrid("Column Set",0,12,"Column Name",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Condition type","%%DYN001-LOW","",DT_WE05_1105_CONDITION_TYPE,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call GetGridContent("",0,"Value of goods rec.",1,"Purchasing Document",DT_WE05_1000_PURCHASING_DOCUMENT,"DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WEWRT_OCC1_OUTPUT")
Call GetGridContent("",0,"Value of goods rec.",2,"Purchasing Document",DT_WE05_1000_PURCHASING_DOCUMENT,"DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WEWRT_OCC1_OUTPUT")
Call VerifyGridCellContent("", 1, "Quantity Received", 0, DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WEMNG)
Call VerifyGridCellContent("", 2, "Quantity Received", 0, DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WEMNG)
Call VerifyGridCellContent("", 1, "Article", 0, DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("", 2, "Article", 0, DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MATNR)
Call VerifyGridCellContent("", 1, "Value of goods rec.", 0, DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WEWRT_OCC1_OUTPUT)
Call VerifyGridCellContent("", 2, "Value of goods rec.", 0,DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WEWRT_OCC1_OUTPUT)
Call VerifyGridCellContent("", 1, "Currency", 0,DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERS)
Call VerifyGridCellContent("", 2, "Currency", 0,DT_WE05_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WAERS)


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


