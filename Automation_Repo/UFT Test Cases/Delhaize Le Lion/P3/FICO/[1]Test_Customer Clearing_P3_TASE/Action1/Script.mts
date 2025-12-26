'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer Clearing_PRE3 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clearing_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-FBL5N----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SelectRadioButton("X_CLSEL","Cleared items", False)           
Call SetTextbox("Clearing date","SO_AUGDT-LOW","",ConvertDate(DT_FBL5N_1000_CLEARING_DATE),False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company Code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call SelectCheckbox("X_SHBV",0, "ON", False)
Call SelectCheckbox("X_MERK",0, "ON", False)
Call SelectCheckbox("X_PARK",0, "ON", False)
Call SelectCheckbox("X_APAR",0, "ON", False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Wait 5
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait 10
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_GET")
'GetRowNo =2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'Call VerifyifGuiLabelExists_ByIndex("S_LEDG",0)
Call TakeScreenShot
Wait 5
Call SelectMenuBar("Settings;Switch list")
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait 5
'''''''''Added 66 to 126 by KGARA on 19/07/2022'''''''''''.
' ''SelectColumnGuiGrid(gridTitle, gridIndex, columnName, blnIsItPopup)
Call SelectColumnGuiGrid("","","Document Number",False)
Call TakeScreenShot
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
'''Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER,True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC0,True)
Call PressEnter()
'''Call VerifyGridCellContentbyName("","1","Document Number","",DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("",1,"Document Number","",DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)

Wait 5
Call SelectMenuBar("Settings;Switch list")
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait 5
'''' ClickLabel(labelContent, labelIndex, blnIsItPopup)
Call ClickLabel("DocumentNo","",False)
Call TakeScreenShot
Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC1,True)
Call PressEnter()
Call VerifyifGuiLabelExists_ByIndex("List does not contain any data",0)

'''Call SelectMenuBar("Settings;Switch list")
'''Call SelectColumnGuiGrid("","","Document Number",False)
Call ClickLabel("DocumentNo","",False)
Call TakeScreenShot
Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC2,True)
Call PressEnter()
Call TakeScreenShot
'''Call VerifyGridCellContentbyName("","1","Document Number","",DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1,0)

'''Call SelectMenuBar("Settings;Switch list")
Call ClickLabel("DocumentNo","",False)
Call TakeScreenShot
Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC3,True)
Call PressEnter()
Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA_OCC1,0)

'''Call SelectMenuBar("Settings;Switch list")
'''Call SelectMenuBar("Settings;Switch list")
Call ClickLabel("DocumentNo","",False)
Call TakeScreenShot
Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC4,True)
Call PressEnter()
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA_OCC2,0)

'''Call SelectMenuBar("Settings;Switch list")
'''Call SelectMenuBar("Settings;Switch list")
Call ClickLabel("DocumentNo","",False)
Call TakeScreenShot
Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC5,True)
Call PressEnter()
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA_OCC3,0)

'''Call SelectMenuBar("Settings;Switch list")
'''Call SelectMenuBar("Settings;Switch list")
Call ClickLabel("DocumentNo","",False)
Call TakeScreenShot
Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC6,True)
Call PressEnter()
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA_OCC4,0)

''''Call SelectMenuBar("Settings;Switch list")
'''Call SelectMenuBar("Settings;Switch list")
Call ClickLabel("DocumentNo","",False)
Call TakeScreenShot
Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC7,True)
Call PressEnter()
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA_OCC5,0)

Call LogOff()
Call FinalStatus ()


'Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
''Call ClickButtonToolBar(buttonName, toolbarIndex)
'Call ClickButtonToolBar("&FIND",0)
''Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call ClickButtonToolBar("&FIND",0)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Clearing Document",True)
''Call SetTextbox("Find","GD_SEARCHSTR","","Clearing Document",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
''Call ClickButtonIfExist("Copy   \(Enter\)",True)
''Call ClickButtonIfExist("Copy   \(Enter\)",True)
''Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST_820",11,True)
''Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST_820",12,True)
''Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButtonIfExist("Define Filter Values",True)
'Call TakeScreenShot
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER,True)
''Call SetTextbox("Clearing Document","%%DYN002-LOW","",DT_FBL5N_1105_CLEARING_DOCUMENT,True)
'Call TakeScreenShot
'Call ClickButtonIfExist("Execute   \(Enter\)",True)
'Call VerifyGridCellContentbyName("shell","","Document Number","",DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
'' VerifyGridCellContentbyName(gridName, gridRowNumber, gridColumnName, gridIndex, expectedValue)
''Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR,2)
'Call TakeScreenShot
'Call SelectMenuBar("Settings;Switch list")
''Call SelectMenuBar("Settings;Switch list")
'Call TakeScreenShot
'Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
'Call ClickButtonIfExist("Find",True)
''Call ClickButtonToolBar("&FIND",0)
''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
''Call ClickButtonIfExist("OK   \(Enter\)",True)
''Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButtonIfExist("Find",True)
''Call ClickButtonToolBar("&FIND",0)
''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Clearing Document",True)
''Call ClickButtonIfExist("OK   \(Enter\)",True)
''Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","Clearing Document",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButtonIfExist("Copy   \(Enter\)",True)
'Call ClickButtonIfExist("Define Filter Values",True)
'Call TakeScreenShot
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC1,True)
''Call SetTextbox("Clearing Document","%%DYN002-LOW","",DT_FBL5N_1105_CLEARING_DOCUMENT,True)
'Call TakeScreenShot
'Call ClickButtonIfExist("Execute   \(Enter\)",True)
''Call VerifyifGuiLabelExists_ByIndex("List does not contain any data",0)
'Call TakeScreenShot
'Call SelectMenuBar("Settings;Switch list")
'Call TakeScreenShot
''Call SelectMenuBar("Settings;Switch list")
'Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
'Call ClickButtonToolBar("&FIND",0)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
'Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
''Call ClickButtonIfExist("Find",True)
''Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
''Call ClickButtonIfExist("Continue   \(Enter\)",True)
''Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButtonToolBar("&FIND",0)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Clearing Document",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
'Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
''Call ClickButtonIfExist("Find",True)
''Call SetTextbox("Find","GD_SEARCHSTR","","Clearing Document",True)
''Call ClickButtonIfExist("Continue   \(Enter\)",True)
''Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
''Call ClickButtonIfExist("Copy   \(Enter\)",True)
'Call ClickButtonIfExist("Define Filter Values",True)
'Call TakeScreenShot
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC2,True)
''Call SetTextbox("Clearing Document","%%DYN002-LOW","",DT_FBL5N_1105_CLEARING_DOCUMENT,True)
'Call TakeScreenShot
'Call ClickButtonIfExist("Execute   \(Enter\)",True)
''Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1,0)
'Call TakeScreenShot
'Call SelectMenuBar("Settings;Switch list")
'Call TakeScreenShot
''Call SelectMenuBar("Settings;Switch list")
'Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
''Call ClickButtonToolBar("&FIND",0)
''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
''Call ClickButtonIfExist("OK   \(Enter\)",True)
''Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call ClickButtonIfExist("Find",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
''Call ClickButtonToolBar("&FIND",0)
''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Clearing Document",True)
''Call ClickButtonIfExist("OK   \(Enter\)",True)
''Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call ClickButtonIfExist("Find",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","Clearing Document",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButtonIfExist("Copy   \(Enter\)",True)
'Call ClickButtonIfExist("Define Filter Values",True)
'Call TakeScreenShot
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC3,True)
''Call SetTextbox("Clearing Document","%%DYN002-LOW","",DT_FBL5N_1105_CLEARING_DOCUMENT,True)
'Call TakeScreenShot
'Call ClickButtonIfExist("Execute   \(Enter\)",True)
''Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA_OCC1,0)
'Call TakeScreenShot
'Call SelectMenuBar("Settings;Switch list")
'Call TakeScreenShot
''Call SelectMenuBar("Settings;Switch list")
'Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
''Call ClickButtonToolBar("&FIND",0)
''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
''Call ClickButtonIfExist("OK   \(Enter\)",True)
''Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call ClickButtonIfExist("Find",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
''Call ClickButtonToolBar("&FIND",0)
''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Clearing Document",True)
''Call ClickButtonIfExist("OK   \(Enter\)",True)
''Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call ClickButtonIfExist("Find",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","Clearing Document",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButtonIfExist("Copy   \(Enter\)",True)
'Call ClickButtonIfExist("Define Filter Values",True)
'Call TakeScreenShot
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC4,True)
''Call SetTextbox("Clearing Document","%%DYN002-LOW","",DT_FBL5N_1105_CLEARING_DOCUMENT,True)
'Call TakeScreenShot
'Call ClickButtonIfExist("Execute   \(Enter\)",True)
''Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA_OCC2,0)
'Call TakeScreenShot
'Call SelectMenuBar("Settings;Switch list")
'Call TakeScreenShot
'''Call SelectMenuBar("Settings;Switch list")
'Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
'Call ClickButtonToolBar("&FIND",0)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
'Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
''Call ClickButtonIfExist("Find",True)
''Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
''Call ClickButtonIfExist("Continue   \(Enter\)",True)
''Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButtonToolBar("&FIND",0)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Clearing Document",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
'Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
''Call ClickButtonIfExist("Find",True)
''Call SetTextbox("Find","GD_SEARCHSTR","","Clearing Document",True)
''Call ClickButtonIfExist("Continue   \(Enter\)",True)
''Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
''Call ClickButtonIfExist("Copy   \(Enter\)",True)
'Call ClickButtonIfExist("Define Filter Values",True)
'Call TakeScreenShot
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC5,True)
''Call SetTextbox("Clearing Document","%%DYN002-LOW","",DT_FBL5N_1105_CLEARING_DOCUMENT,True)
'Call TakeScreenShot
'Call ClickButtonIfExist("Execute   \(Enter\)",True)
''Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA_OCC3,0)
'Call TakeScreenShot
'Call SelectMenuBar("Settings;Switch list")
'Call TakeScreenShot
''Call SelectMenuBar("Settings;Switch list")
'Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
''Call ClickButtonToolBar("&FIND",0)
''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
''Call ClickButtonIfExist("OK   \(Enter\)",True)
''Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call ClickButtonIfExist("Find",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
''Call ClickButtonToolBar("&FIND",0)
''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Clearing Document",True)
''Call ClickButtonIfExist("OK   \(Enter\)",True)
''Call ClickButtonIfExist("Cancel   \(F12\)",True)
''Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call ClickButtonIfExist("Find",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","Clearing Document",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButtonIfExist("Copy   \(Enter\)",True)
'Call ClickButtonIfExist("Define Filter Values",True)
'Call TakeScreenShot
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBL5N_1105_DOCUMENT_NUMBER_OCC7,True)
''Call SetTextbox("Clearing Document","%%DYN002-LOW","",DT_FBL5N_1105_CLEARING_DOCUMENT,True)
'Call TakeScreenShot
'Call ClickButtonIfExist("Execute   \(Enter\)",True)
''Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_LIST_CONTAINS_NO_DATA_OCC5,0)
'Call TakeScreenShot
'Call SelectMenuBar("Settings;Switch list")
'Call TakeScreenShot
''------------------------'Log Off From Applicaton--------------------------------
'Call LogOff()
'Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

