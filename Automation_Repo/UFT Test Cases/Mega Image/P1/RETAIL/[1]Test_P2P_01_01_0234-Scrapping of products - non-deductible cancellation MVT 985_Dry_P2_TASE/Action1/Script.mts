
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0273-Return process in DC_P1
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0234-Scrapping of products - non-deductible cancellation MVT 985_Dry_P2
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_0234- MVT 985_Dry_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_P2P_01_01_0234-Scrapping of products - non-deductible cancellation MVT 985_Dry_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot()

'
Call  SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_2010_GODYNPROMAT_DOC, False)


Call TakeScreenShot()

Call PressEnter()

Call PressEnter()
Call TakeScreenShot()
Call ClickButtonIfExist("Open detail data",False)
Call SelectTab("TS_GOITEM", "Where", False)

Call SetTextBox("Reason for Movement","GOITEM-GRUND",0,DT_MIGO_0325_REASON_FOR_MOVEMENT,False)

Call PressEnter()

Call ClickButton("OK_NEXT_ITEM",False)

Call SetTextBox("Reason for Movement","GOITEM-GRUND",0,DT_MIGO_0325_REASON_FOR_MOVEMENT,False)

Call PressEnter()

Call ClickButton("OK_NEXT_ITEM",False)

Call SetTextBox("Reason for Movement","GOITEM-GRUND",0,DT_MIGO_0325_REASON_FOR_MOVEMENT,False)

Call PressEnter()

Call ClickButton("OK_NEXT_ITEM",False)

Call SetTextBox("Reason for Movement","GOITEM-GRUND",0,DT_MIGO_0325_REASON_FOR_MOVEMENT,False)

Call PressEnter()

Call ClickButton("OK_NEXT_ITEM",False)

Call SetTextBox("Reason for Movement","GOITEM-GRUND",0,DT_MIGO_0325_REASON_FOR_MOVEMENT,False)

Call PressEnter()

Call ClickButton("OK_NEXT_ITEM",False)

Call SetTextBox("Reason for Movement","GOITEM-GRUND",0,DT_MIGO_0325_REASON_FOR_MOVEMENT,False)

Call PressEnter()'

Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)

Call CLickButton("Check Entries   \(F7\)",False)

Call SetTableData("SAPLMIGOTV_GOITEM","OK",1,"","","ON",False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK",2,"","","ON",False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK",3,"","","ON",False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK",4,"","","ON",False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK",5,"","","ON",False)



Call TakeScreenShot()

Call ClickButton("Post Document   \(Shift\+F11\)",False)

Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

Call TakeScreenShot()

Call SetTcode(DT_MIGO_0001_OKCD)

Call PressEnter()

Call SetTextBox("Article Doc\.","RM07M-MBLNR",0,DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)

Call SetTextBox("Art\. Doc\. Year","RM07M-MJAHR",0,DT_MB90_0460_ART_DOC_YEAR,False)

Call TakeScreenShot()

Call PressEnter()
WAIT 5
Call PressEnter()

Call ClickButton("Accounting Documents   \(F7\)",False)


Call GetGridContentByRefColumn("Documents in Accounting", 0, "Object type text", "Accounting Document"," Doc. Number","DT_MIGO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT")



Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting Document", False)

Call ClickButton("Display Document   \(F2\)",False)



Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)

Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("", 4, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)


Call VerifyGridCellContent("", 5, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)

Call VerifyGridCellContent("", 6, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)

Call VerifyGridCellContent("", 7, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BSCHL)

Call VerifyGridCellContent("", 8, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_7_BSCHL)

Call VerifyGridCellContent("", 9, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_8_BSCHL)

Call VerifyGridCellContent("", 10, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_9_BSCHL)

Call VerifyGridCellContent("", 11, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_10_BSCHL)




Call VerifyGridCellContent("", 1, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XNEGP)

Call VerifyGridCellContent("", 2, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XNEGP)

Call VerifyGridCellContent("", 3, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_XNEGP)

Call VerifyGridCellContent("", 4, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_XNEGP)

Call VerifyGridCellContent("", 5, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_XNEGP)

Call VerifyGridCellContent("", 6, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_XNEGP)

Call VerifyGridCellContent("", 7, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_XNEGP)

Call VerifyGridCellContent("", 8, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_7_XNEGP)

Call VerifyGridCellContent("", 9, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_8_XNEGP)

Call VerifyGridCellContent("", 10, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_9_XNEGP)

Call VerifyGridCellContent("", 11, "Negative posting", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_10_XNEGP)



Call VerifyGridCellContent("", 1, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)

Call VerifyGridCellContent("", 2, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 3, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("", 4, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

Call VerifyGridCellContent("", 5, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)


Call VerifyGridCellContent("", 6, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)

Call VerifyGridCellContent("", 7, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_KTONR)

Call VerifyGridCellContent("", 8, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_7_KTONR)

Call VerifyGridCellContent("", 9, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_8_KTONR)

Call VerifyGridCellContent("", 10, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_9_KTONR)

Call VerifyGridCellContent("", 11, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_10_KTONR)

Call TakeScreenShot()
''
Call SetTcode(DT_MIGO_0750_OKCD)

Call PressEnter()

Call TakeScreenShot()

'Call CheckTcodeScreen(DT_MIGO_0750_OKCD)

Call SetTextBox("Company Code","BUKRS-LOW",0,DT_MIGO_1000_COMPANY_CODE,False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextBox("Article Document","MBLNR-LOW",0,DT_MIGO_1000_ARTICLE_DOCUMENT,False)

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Set Filter   (Ctrl+Shift+F2)",False)
''Call SelectCellGuiGrid("Column Set",0,6,"Column Name",True)
Call SelectCellGuiGridByCellValue("Column Set",0,"Column Name","Movement type",False)

Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Movement type","%%DYN001-LOW","",DT_MIGO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART,True)
Call ClickButton("Execute   \(Enter\)",True)


Call VerifyGridCellContent("", 1, "Movement type", 0, DT_MIGO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART)

Call LogOff()

Call FinalStatus()

Public Function SelectCellGuiGridByCellValue(gridTitle,gridIndex,refColumn,refFieldVal,blnIsItPopup)
If Not (Environment.Value("blnFatalError")) Then
	   If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : SelectCellGuiGridByCellValue"
Dim objGridGuiComp80,objGridGuiComp201,objGrid

Set objWindow= SetSAPwindowObj(blnIsItPopup)
strStepName = "Select a cell in GuiGrid" 
If refColumn<> ""  then
Set objGridGuiComp80 = sapguisession(sessionObject).sapguiwindow(objWindow).SAPGuiGrid("guicomponenttype:=80","name := shell.*","title:="&gridTitle,"index:="&gridIndex)
Set objGridGuiComp201= sapguisession(sessionObject).sapguiwindow(objWindow).SAPGuiGrid("guicomponenttype:=201","name:=shell.*","title:="&gridTitle,"index:="&gridIndex)
Set objGrid=SetObj(objGridGuiComp80,objGridGuiComp201)
		If objGrid.Exist Then
			On error resume Next
					strRowNo=objGrid.FindRowByCellContent(refColumn,refFieldVal)
					strRowNumber=strRowNo
            		If Err.Number=0 and strRowNumber <> "" Then
						objGrid.SelectCell strRowNumber,refColumn
						If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
				  				ImagePath=CaptureScreenshot(strStepName,objGrid,False,False,True)
	          			End if
						Call ReporterFunction(strLibraryFileName,"SelectCellGuiGridByCellValue","2","SelectCellGuiGridByCellValue","Cell Selected:-"& vbNewLine &" Row No : "&strRowNumber &" Column : "&refColumn)
						strMsg="Cell of  Row [" & strRowNumber & "] and column [" & refColumn & "] selected in GuiGrid."
						strStatus = "DONE"
					ElseIf Err.Number=-2147220983 Then
						Call ReporterFunction(strLibraryFileName,"SelectCellGuiGridByCellValue","1","SelectCellGuiGridByCellValue","The Column:  "&refColumn& " Not Found Within The GuiGrid")
						 strStatus = "FAIL"
						 strMsg = "Column ["&refColumn& "] Not Found Within The GuiGrid"
						 blnObjectError=True
'					Else 
'						blnObjectError=True
'						Call ReporterFunction(strLibraryFileName,"SelectRowGuiGrid","1","SelectRowGuiGrid","Value "&refValue&" Not Found Within The Column :"&refColumn& " of GuiGrid")
'						strStatus = "FAIL"
'						strMsg ="Value ["&refValue& "] Not Found Within The Column [" & refColumn & "] of GuiGrid"
						
					End If
		Else 
			blnObjectError=True
			Call ReporterFunction(strLibraryFileName,"SelectCellGuiGridByCellValue","1","SelectCellGuiGridByCellValue","GuiGrid Object Not Found,Please verify the Screen")
		    strStatus = "FAIL"
			strMsg = "GuiGrid Object Not Found,Please verify the Screen"
			
		End If
Else
	Call ReporterFunction(strLibraryFileName,"SelectCellGuiGridByCellValue","1","SelectCellGuiGridByCellValue","Function Parameter Not Passed Properly. Check the --SelectCellGuiGridByCellValue-- Function Call")
	strStatus = "FAIL"
	strMsg = "Function Parameter Not Passed Properly. Check the --SelectCellGuiGridByCellValue-- Function Call"
End if				
       If  blnObjectError  Then
		    Environment.Value("blnFatalError")=True
	   End If

	   If strStatus = "FAIL"  Then
		    SelectCellGuiGridByCellValue = strMsg
		    blnMainFailFlag = True
		    ImagePath=CaptureScreenshot(strStepName,objGrid,False,False,True)
	   Else
		    SelectCellGuiGridByCellValue = strMsg
	   End If
	
	   If blnDefault_eSwiftReporting Then  
		   Call UpdateResultHtml(strStepName,refValue,strMsg,strStatus,"")
	   End If
 ' //////////Word Document ///////////
'				If blnCreateTrainingDoc Then
'					Check=CheckStatus_Doc(strStatus)
'					If Check=1 Then
'						Call storeTableorGridCoordinates(objGrid)
'						Call StoreSteps(strStepName)
'					End If
'				End If
 ' //////////Word Document ///////////
Set objGrid=Nothing
Set objWindow=Nothing
Set objGridGuiComp80=Nothing
Set objGridGuiComp201=Nothing
End If
End Function

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





