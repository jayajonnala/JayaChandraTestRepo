
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Order Proposal (DLI150)_Get_IDoc_TASE
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

gstrTestCaseName = "Test_Order Proposal (DLI150)_Get_IDoc_TASE"
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''' Login '''
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)



Call SetTextboxNoLabel("CRETIM-LOW","",DT_WE02_1100_CREATED_AT,False)
''Call SetTextbox("to","CRETIM-HIGH","",DT_WE02_1100_CREATED_AT_TO,False)
Call SetTextboxNoLabel("CREDAT-LOW","",ConvertDate(DT_WE02_1100_CREATED_ON),False)
Call SetTextboxNoLabel("CREDAT-HIGH","",ConvertDate(DT_WE02_1100_TO),False)
Call SetTextboxNoLabel("IDOCTP-LOW","",DT_WE02_1100_BASIC_TYPE,False)

Call SetTextboxNoLabel("MESTYP-LOW","",DT_WE02_1100_LOGICAL_MESSAGE,False)

'Call SetTextboxNoLabel("Created On","CREDAT-LOW","",ConvertDate(DT_WE02_1100_CREATED_ON),False)
'Call SetTextboxNoLabel("to","CREDAT-HIGH","",ConvertDate(DT_WE02_1100_TO),False)
'Call SetTextboxNoLabel("Basic Type","IDOCTP-LOW","",DT_WE02_1100_BASIC_TYPE,False)

'Call SetTextboxNoLabel("Logical Message","MESTYP-LOW","",DT_WE02_1100_LOGICAL_MESSAGE,False)


Call SetTextboxNoLabel("MESCOD-LOW",0,DT_DLI,False)
Call SetTextboxNoLabel("MESFCT-LOW",0,DT_150,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
''***************** Added if loop as this below objects exist sometimes
If VerifyColumnGuiGrid("shell",0,"Created At")  Then
	'Call SelectColumnGuiGrid("Selected IDocs",0,"IDoc Number",False)
	Call SelectColumnGuiGrid("Selected IDocs",0,"Created At",False)
	Call ClickButtonToolBar("&SORT_DSC",0)
	''ActivateCellGuiGridByRefVal(gridTitle,gridIndex,refColumn,refFieldVal,columnName,blnIsItPopup)
	Call SelectRowGuiGridbyRowNo("Selected IDocs",0,1,False)
	
	''Call SelectAllRowGuiGrid("Selected IDocs",0,False)
	''Call SelectCellGuiGrid("Selected IDocs",0,1,"Segments",False)
	Call ActivateCellGuiGridByRefVal("Selected IDocs",0,"Segments","18","IDoc Number",False)
End If
'**************************
Call VerifyTextBoxContent("Current Status","EDIDC-STATUS",0,DT_STATUS,False)

Call ActivateItemGuiTree(0,"#1;#3;#1","#1")

Call TakeScreenShot()

Call SelectTab("TABSTRIP_EDIDS", "Sts details", False)
Call GetTextboxValue("EDI_INTDS-DOCNUM",0,"DT_IDOC_NUMBER_OUTPUT",False)

''DT_EXPECTEDVALUE
Call GetTextboxValue("EDI_INTDS-STAPA2",0,"DT_PO_NUMBER_OUTPUT",False)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()

