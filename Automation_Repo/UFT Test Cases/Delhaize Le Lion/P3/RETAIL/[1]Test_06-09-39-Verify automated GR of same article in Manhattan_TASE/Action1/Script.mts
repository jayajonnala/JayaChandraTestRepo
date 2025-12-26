
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06-09-14-correction on reception (WMS code 6)_P2_checkinSAP_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_06-09-39-Verify article Manhattan"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-01-01-01-01-05-Send PO to WMS_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''--------TransactionCode-ME23N----------''''
 
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Call ClickButton("Display/Change   \(F7\)",False)
'Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME23N_1105_MEPO_TOPLINEBSART)
Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME21N_0003_PUR_ORDER,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call TakeScreenShot
If IfTabExists("ITEM_DETAIL","Purchase Order History",False) <> True Then
'Call SetTcode(DT_T_CODE) 
'Call PressEnter()     
'Call TakeScreenShot
'Call SetTextbox("Created On","CREDAT-LOW","",DT_DATE,False)
'Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE02_1100_MESSAGE_VARIANT,False)
'Call SetTextbox("Message Function","MESFCT-LOW","",DT_WE02_1100_MESSAGE_FUNCTION,False)
'''Call SetTextbox("Created At","CRETIM-LOW","",DT_GR_TIME,False)
'''Call SetTextbox("Last Changed At","UPDTIM-LOW","",DT_GR_TIME,False)
'Call PressEnter() 
'Call TakeScreenShot
'Call ClickButton("Execute   \(F8\)",false)
'Call TakeScreenShot
''''Call SelectColumnGuiGrid("Selected IDocs", 0,  "Created at", False)
''''Call ClickButtonToolbar("&SORT_DSC",0)
''''Call SelectRowGuiGridbyRowNo("Selected IDocs", 0,1,False)
''''Call DoubleClickGuiGridCell("Selected IDocs", 0, 1, "IDoc number", False)
''''Call TakeScreenShot
''''''Call ClickLinkGuiTree("#1;#2;#1;#1","Segment 000002",0,False)
'''outputVal=VerifyTableCellContent(4, "Fld Cont.", "IDOC_TREE_CONTROLINT_SEG_CONTROL", DT_ME21N_0003_PUR_ORDER) 
'''rowNum=1
'''While outputVal<>True
'''	rowNum=rowNum+1
'''	Call ClickButton("Back   \(F3\)",false)
'''	''Call DoubleClickGuiGridCell("Selected IDocs", 0, rowNum, "IDoc number", False)
'''	Call SelectRowGuiGridbyRowNo("Selected IDocs", 0,rowNum,False)
'''	Call DoubleClickGuiGridCell("Selected IDocs", 0, rowNum, "IDoc number", False)
'''	Call TakeScreenShot
''Call ActivateItemGuiTree(0, "#1;#3;#2", "50")
'Call ActivateItemGuiTree(0, "#1;#3;#4", "50")
'Call TakeScreenShot
'
'''outputVal=VerifyTableCellContent(4, "Fld Cont.", "IDOC_TREE_CONTROLINT_SEG_CONTROL", DT_ME21N_0003_PUR_ORDER) 
'''Wend
'''''Call SelectRowGuiGridbyRowNo("Selected IDocs", 0,1,False)
''Call ClickButton("Back   \(F3\)",false)
'Call GetTextboxValue("EDI_INTDS-DOCNUM", 0, "DT_IDOC_NO_OUTPUT", False)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call SetTcode(DT_T_CODE_1) 
'Call PressEnter()     
'Call TakeScreenShot
'
'Call SetTextbox("IDoc Number", "SX_DOCNU-LOW", 0, DT_IDOC_NO, False)
'Call SetTextbox("Created On", "SX_CREDA-LOW", 0, DT_DATE, False)
'Call SetTextbox("Changed On", "SX_UPDDA-LOW", 0, DT_DATE, False)
'Call TakeScreenShot
'Call ClickButton("Execute   \(F8\)",false)
'Call TakeScreenShot
''Call ClickLinkGuiTree("Retail Pre-Production;IDoc in inbound processing;IDoc ready to be passed to application", "IDoc ready to be passed to application", 0, False)
'Call ClickLinkGuiTree("Retail Pre-Production;IDoc in inbound processing;Application document posted;MBGMCR", "MBGMCR", 0, False)
'
'Call ClickButton("Process Selected Node   \(F8\)", False)
''''''Call ActivateItemGuiTree(0, "Retail Pre-Production;IDoc in inbound processing;IDoc ready to be passed to application", "IDoc ready to be passed to application")
'Call VerifyGridCellContent("Processed IDocs", 1, "new status", 0, DT_IDOC_STATUS)
'Call SetTcode(DT_SAPTRANSACTIONCODE) 
'Call PressEnter()     ' 
'
'Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
'Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME21N_0003_PUR_ORDER,True) 
'Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
'Call ClickButton("Other Document   \(Enter\)",True)
'Call TakeScreenShot
'Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
'Call TakeScreenShot
'End If
'
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call TakeScreenShot
End	If
'''Call FindRowNumber("","Movement type",DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART,"DT_ROW")
'''GetRowNo =2
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'''Call ClickCellGuiGrid("","","Article Document",DT_ROW,"Movement type",DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART,False)
'''
''Call DoubleClick()
''Call TakeScreenShot
'''Call VerifyTextBoxContent("HeaderText","GOHEAD-BKTXT",0,DT_CHECK_HEADERTEXT,False)
'''Call VerifyTextBoxContent("Movement type","GOITEM-BWART",0,DT_CHECK_MVTTYPE,False)

Call LogOff()
Call FinalStatus ()

Public Function VerifyIfTabExists(tabStripName,tabName,blnIsItPopup)
	If Not (Environment.Value("blnFatalError")) Then
	   If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : VerifyIfTabExists"
	Dim objTabStrip,objWindow
	Set objWindow= SetSAPwindowObj(blnIsItPopup)
	strStepName = "Verify if tab exist: " & tabName 
	If tabStripName <> "" and tabName <> "" Then
     	Set objTabStrip1 = sapguisession(sessionObject).sapguiwindow(objWindow).SAPGuiTabStrip(guiTab,"name:="&tabStripName)
		Set objTabStrip2 = sapguisession(sessionObject).sapguiwindow(objWindow).SAPGuiTabStrip("name:="&tabStripName) 'to identify child tabstrips
		Set objTabStrip = SetObj(objTabStrip1,objTabStrip2)
		If  objTabStrip.Exist   Then
			strAllItems = objTabStrip.GetROProperty("allitems")
			strAllItems = ";"&strAllItems
			strTabNameAppended =";"&tabName&";" 
			If instr(1,strAllItems,strTabNameAppended,1) Then
				''objTabStrip.Select tabName
				If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
				  ImagePath=CaptureScreenshot(strStepName,objTabStrip,False,False,True)
	          End if
				Call ReporterFunction(strLibraryFileName,"VerifyIfTabExists","2","Tab Field","Tab Exists :-"& vbNewLine &" Tab Name : "&tabName)
				strStatus="DONE"
				strMsg="Tab Exists"
			Else
				Call ReporterFunction(strLibraryFileName,"VerifyIfTabExists","1","Tab Field","Tab Not Found Within The Tab Strip . Please Verify --VerifyIfTabExists-- Function Call")
				strStatus = "FAIL"
				strMsg = " Tab Not Found Within The Tab Strip. Please Verify --VerifyIfTabExists-- Function Call"
				'blnObjectError=True	
			End If
		Else 
			Call ReporterFunction(strLibraryFileName,"VerifyIfTabExists","1","Tab Field","Tab Strip Not Found. Please Verify --VerifyIfTabExists-- Function Call")
			 strStatus = "FAIL"
			 strMsg = " Tab Strip Not Found. Please Verify --VerifyIfTabExists-- Function Call"
			 blnObjectError=True
		End If
	Else 
		Call ReporterFunction(strLibraryFileName,"VerifyIfTabExists","1","Tab Field","Function Parameter Not Passed Properly. Check the --VerifyIfTabExists-- Function Call")
		strStatus="FAIL"
		strMsg = "Function Parameter Not Passed Properly. Check the --VerifyIfTabExists-- Function Call"
	End If	
	If  blnObjectError  Then
		    Environment.Value("blnFatalError")=True
	   End If

	   If strStatus = "FAIL"  Then
		    VerifyIfTabExists = strMsg
		    blnMainFailFlag = True
		    ImagePath=CaptureScreenshot(strStepName,objTabStrip,False,False,True)
	   Else
		    VerifyIfTabExists = True
	   End If
	
	   If blnDefault_eSwiftReporting Then  
		   Call UpdateResultHtml(strStepName,"",strMsg,strStatus,"")
	   End If
	Set objTabStrip = nothing

Set objTabStrip=Nothing
Set objWindow=Nothing
End If
End Function

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



