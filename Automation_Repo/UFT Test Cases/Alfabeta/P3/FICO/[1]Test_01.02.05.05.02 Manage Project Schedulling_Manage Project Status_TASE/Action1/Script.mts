'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01.02.05.05.03 Manage Project Schedulling_Manage Project Status_
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_01.02.05.05.03 Manage Project Schedulling_Manage Project Status_"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-CJ20N ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INC_VAL",Cint(DT_INC_VAL)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call Click202ButtonToolBar("WK_CP",1)
Call TakeScreenShot
Call ActivateNodeGuiTree(1, "#1;#1")
Call TakeScreenShot

Call SetTextbox("Project def\.","PROJ-PSPID", 0, DT_CJ20N_3990_PROJECT_DEF, False)
Call SetTextbox("Project def\.","PROJ-POST1", 0, DT_CJ20N_3990_PROJECT_DEF_OCC1, False)
Call SetComboByKey("Project Profile", DT_CJ20N_1404_PROJECT_PROFILE)
Call TakeScreenShot
Call  SelectTab("PTABSCR", "Basic Data", False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call SetTextbox("Company code","PROJ-VBUKR", 0, DT_CJ20N_1205_COMPANY_CODE, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait 2
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)


Call Click202ButtonToolBar("OPEN",0)
Call SetTextbox("Project definition","CNPB_W_ADD_OBJ_DYN-PROJ_EXT", 0, DT_CJ20N_0900_PROJECT_DEFINITION, True)
Call ClickButton("Open   \(Enter\)",True)

Call Click202ButtonToolBar("WBSE_OVW",3)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)

Call SetTableData("SAPLCJWBTAB_902", "Level", 2, "", "", DT_CJ20N_0902_TABLECELL_LEVEL_1, False)

Call SetTableData("SAPLCJWBTAB_902", "WBS element", 1, "", "", DT_CJ20N_0902_TABLECELL_WBS_ELEMENT_0, False)
Call SetTableData("SAPLCJWBTAB_902", "WBS element", 2, "", "", DT_CJ20N_0902_TABLECELL_WBS_ELEMENT_1, False)

Call SetTableData("SAPLCJWBTAB_902", "Description", 1, "", "", DT_CJ20N_0902_TABLECELL_DESCRIPTION_0, False)
Call SetTableData("SAPLCJWBTAB_902", "Description", 2, "", "", DT_CJ20N_0902_TABLECELL_DESCRIPTION_1, False)

Call  SelectTab("TABCJLE", "Control", False)
Call TakeScreenShot

Call  SelectTab("TABCJLE", "Responsibilities", False)
Call TakeScreenShot

Call SetTableData("SAPLCJWBTAB_904", "Req.cost center", 1, "", "", DT_CJ20N_0904_TABLECELL_REQCOST_CENTER_0, False)
Call SetTableData("SAPLCJWBTAB_904", "Req.cost center", 2, "", "", DT_CJ20N_0904_TABLECELL_REQCOST_CENTER_1, False)

''Call  SelectTab("TABCJLE", "Basic Data", False)
Call PressEnter()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait 2
Call ClickButton("Confirm   \(Enter\)",True)
Call TakeScreenShot

Call SetComboByKey("Proj\.type", DT_CJ20N_0902_TABLECELL_TYP_0)
Call PressEnter()
Call ClickButton("Confirm   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Investment program","RAIP1-PRNAM", 0, DT_CJ20N_0620_INVESTMENT_PROGRAM, True)
Call SetTextboxNoLabel("RAIP1-GJAHR", 0, YEar(DT_CJ20N_0620_RAIP1GJAHR), True)
Call SetTextbox("Position ID","RAIP1-POSID", 0, DT_CJ20N_0620_POSITION_ID, True)

Call ClickButton("Continue   \(Enter\)",True)
Call VerifyStatusBarMessageType("S")

Call Click202ButtonToolBar("OPEN",0)
Call SetTextbox("Project definition","CNPB_W_ADD_OBJ_DYN-PROJ_EXT", 0, DT_CJ20N_0900_PROJECT_DEFINITION_OCC1, True)
Call ClickButton("Open   \(Enter\)",True)

Call SelectMenuBar("Edit;Status;Release")
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)

Call VerifyTextBoxContent("System Status","CNJ_STAT-STTXT_INT", 0, DT_CJ20N_0700_CHECK_TEXT_OF_SYSTEM_STATUS, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC6)

Call Click202ButtonToolBar("OPEN",0)
Call SetTextbox("Project definition","CNPB_W_ADD_OBJ_DYN-PROJ_EXT", 0, DT_CJ20N_0900_PROJECT_DEFINITION_OCC2, True)
Call ClickButton("Open   \(Enter\)",True)

Call VerifyTextBoxContent("System Status","CNJ_STAT-STTXT_INT", 0, DT_CJ20N_0700_CHECK_TEXT_OF_SYSTEM_STATUS, False)
Call SelectMenuBar("Edit;Status;Close;Set")
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)
Call VerifyTextBoxContent("System Status","CNJ_STAT-STTXT_INT", 0, DT_CJ20N_0700_CHECK_TEXT_OF_SYSTEM_STATUS_OCC1, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)
Call TakeScreenShot

Call Click202ButtonToolBar("OPEN",0)
Call SetTextbox("Project definition","CNPB_W_ADD_OBJ_DYN-PROJ_EXT", 0, DT_CJ20N_0900_PROJECT_DEFINITION_OCC3, True)
Call ClickButton("Open   \(Enter\)",True)
Call VerifyTextBoxContent("System Status","CNJ_STAT-STTXT_INT", 0, DT_CJ20N_0700_CHECK_TEXT_OF_SYSTEM_STATUS_OCC1, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call LogOff()
Call FinalStatus()



































Function Click202ButtonToolBar(buttonName,toolbarIndex)

     If Not Environment.Value("blnFatalError") Then
					
				
					If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : Click202ButtonToolBar"
					strFuncName = "Click202ButtonToolBar"
					Val = "<b>Button Name  :</b>" & buttonName 
					strStepName="Click Button " & "'" &buttonName&"'" &"in Tool bar"
					Set objWindow= SetSAPwindowObj(blnIsItPopup)
							  If buttonName <> ""  Then
'											Set objToolbar202=SAPGuisession(sessionObject).sapguiwindow(windowObject).SAPGuiToolbar("guicomponenttype:=202","name:=shel.*","index:="&toolbarIndex)
											Set objToolbar=SAPGuisession(sessionObject).sapguiwindow(windowObject).SAPGuiToolbar("guicomponenttype:=202","name:=shel.*","index:="&toolbarIndex)
'											Set objToolbar=SetObj(objToolbar202,objToolbar202)
															If objToolbar.Exist Then
'																objToolbar.highlight			
																		   On Error Resume Next
																							  If VerifyObjectStatus(objToolbar) Then
																												Call storeCoordinates(objToolbar)
																												If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
																																ImagePath=CaptureScreenshot(strStepName,objToolbar,True,True,True)
																												End If
'																												msgbox Err.Number
																												If Err.Number =0 Then
																													
																											
																												objToolbar.PressButton buttonName
																												strStatus = "DONE"
																												strMsg = "Button Name "&buttonName&" Pressed "
																												Call ReporterFunction(strLibraryFileName,strFuncName,"2",strStepName,strMsg)
																												
																							  ElseIf Err.Number = 424 Then
																							  					strStatus = "FAIL"
																												strMsg = "The Button "&buttonName&" in the toolbar  is not Found."
																												blnObjectError=True
																												Call ReporterFunction(strLibraryFileName,strFuncName,"1",strStepName,strMsg)
																							  
																							  Else
																												strStatus = "FAIL"
																												strMsg = "Toolbar is Disabled, So Button can't be clicked"
																												blnObjectError=True
																												Call ReporterFunction(strLibraryFileName,strFuncName,"1",strStepName,strMsg)
																							  End If
																							  	End If
																Else
																			strStatus = "FAIL"
																			strMsg = " The  toolbar is not Found, Please verify the Screen"
																			blnObjectError=True
																			Call ReporterFunction(strLibraryFileName,strFuncName,"1",strStepName,strMsg)
																			
																 End If
								Else
											strStatus = "FAIL"
											strMsg = "Function Parameter Not Passed Properly. Check the --Click202ButtonToolBar-- Function Call"                                         				
											Call ReporterFunction(strLibraryFileName,strFuncName,"1",strStepName,strMsg)

								End If	
								If  blnObjectError  Then
																		Environment.Value("blnFatalError")=True
																	End If
										
																	If strStatus = "FAIL"  Then
																		Click202ButtonToolBar = strMsg
																		blnMainFailFlag = True
																		ImagePath=CaptureScreenshot(strStepName,objToolbar,True,True,True)
																	Else
																		Click202ButtonToolBar = True
																	End If
										
																	If blnDefault_eSwiftReporting Then  
																			Call UpdateResultHtml (strStepName,Val,strMsg,strStatus,"")																	
																	End If
																	Set objToolbar=nothing
'																	Set objToolbar202=nothing
'																	Set objToolbar202=nothing
					
     End If
		

End Function
'************************************************************************************
'End Function - Click202ButtonToolBar
'************************************************************************************

