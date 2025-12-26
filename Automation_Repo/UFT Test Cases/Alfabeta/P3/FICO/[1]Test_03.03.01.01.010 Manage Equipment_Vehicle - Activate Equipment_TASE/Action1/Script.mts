

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.01.010 Manage Equipment_Vehicle - Activate Equipment
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

gstrTestCaseName = "Test_03.03.01.01.010 Manage Equipment_Vehicle - Activate Equipment"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------Tcode IE05----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter Details
Call SetTextbox("Equipment","EQUNR-LOW","",DT_IE05_1000_EQUIPMENT,False) 
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Yes",True)
wait(2)

'Click on Status Button
Call ClickButton("%_AUTOTEXT003",False)
Call TakeScreenShot()

Call SelectRadioButtonTableByRefColumn("SAPLBSVATC_E","X","Status","ACTV","ON")
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'Click on "Equipment;Functions;Active <-> Inactive;Activate"
Call SelectMenuBar("Equipment;Functions;Active <-> Inactive;Activate")
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_IE05_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Equipment "&DT_IE05_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" changed" )

''----------------------Tcode IE03----------------------------
Call SetTcode(DT_IE05_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_IE05_1000_OKCD)

'Enter Details
Call SetTextbox("Equipment","RM63E-EQUNR","",DT_IE05_1000_EQUIPMENT,False) 
Call TakeScreenShot()
Call PressEnter()
Call PressEnter() 

'Verify equipment Status
Call VerifyTextBoxContent("Status","ITOBATTR-STTXT",0,DT_IE05_1526_CHECK_TEXT_OF_STATUS_OCC1,False)
Call VerifyTextBoxContent("Status","ITOBATTR-STTXU",0,DT_IE05_1526_CHECK_TEXT_OF_STATUS,False)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************



Public Function SelectRadioButtonTableByRefColumn(tableName,columnName,refColumnName,refCellValue,OnOffStatus)
   If Not (Environment.Value("blnFatalError")) Then
	   If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : SelectRadioButtonTableByRefColumn"
   Dim objTable
	'OnOffStatus=UCase(InputValidation(OnOffStatus))
	strStepName = "Select RadioButton With In The Table Cell"
	If tableName<> "" and columnName<>""and refColumnName<>"" and (UCase(OnOffStatus)="ON"  or Ucase(OnOffStatus)="OFF" )Then
		If OnOffStatus="OFF" Then
			strStepName="Deselect RadioButton With In The Table Cell"
		End If
		'OnOffStatus=InputValidation(OnOffStatus)
		Set objTable= SAPGuisession(sessionObject).sapguiwindow(windowObject).sapguitable(guiTable,"name:="&tableName)
		If objTable.Exist Then
			On Error Resume Next
			strRowNumber = objTable.FindRowByCellContent(refColumnName,refCellValue)
			'msgbox err.number
			If Err.Number = 0 Then
				Err.Clear
				On Error Resume Next
				objTable.SetCellData strRowNumber,columnName,OnOffStatus
				If Err.Number=0  Then
				    IF objTable.IsCellEditable(strRowNumber,columnName)=False Then
				    	If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
						   ImagePath=CaptureScreenshot(strStepName,objTable,False,False,False)
					    End if
						Call ReporterFunction(strLibraryFileName,"SelectRadioButtonTableByRefColumn","2","Table Data Field","Value Used :-" & vbNewLine &" Table Name:  '"& tableName&"'"&vbNewLine&"  Column Field:   '"& columnName & "'"&vbNewLine &"   Row Field:  '"&strRowNumber&"'"& vbNewLine &"    ValueField: '"&OnOffStatus&"'")
						 strStatus = "DONE"
						 strMsg="Radio Button from the table with column as reference is selected successfully"
						
					Else
						Call ReporterFunction(strLibraryFileName,"SelectRadioButtonTableByRefColumn","1","Table Data Field"," Table Cell Not Editable. Check the --SelectRadioButtonTableByRefColumn-- Function Call")
						strMsg = "Table Cell Not Editable. Check the --SelectRadioButtonTableByRefColumn-- Function Call"
						strStatus = "FAIL"
						blnObjectError=True					
						
					End If 					 
			    ElseIf Err.Number=-2147220983 Then
						Call ReporterFunction(strLibraryFileName,"SelectRadioButtonTableByRefColumn","1","Table Data Field","Column Not Found. Check the --columnName--Parameter of--SelectRadioButtonTableByRefColumn-- Function Call")
						strMsg = "Column Not Found. Check the --columnName--Parameter of --SelectRadioButtonTableByRefColumn-- Function Call"
						strStatus = "FAIL"
						blnObjectError=True
				ElseIF Err.Number = 13 Then
					Call ReporterFunction(strLibraryFileName,"SelectRadioButtonTableByRefColumn","1","Grid Data Field","refCellValue Not Found Within RefColumn. Check the --refCellValue--Parameter of--SelectRadioButtonTableByRefColumn-- Function Call")
					strMsg = "refCellValue Not Found Within RefColumn. Check the --refCellValue--Parameter of --SelectRadioButtonTableByRefColumn-- Function Call"
					strStatus = "FAIL"
					blnObjectError=True	
			    End If
			ElseIf Err.Number=-2147220983 Then
					Call ReporterFunction(strLibraryFileName,"SelectRadioButtonTableByRefColumn","1","Table Data Field","Column Not Found. Check the --columnName--Parameter of--SelectRadioButtonTableByRefColumn-- Function Call")
					strMsg = "refColumn Not Found. Check the --refColumnName--Parameter of --SelectRadioButtonTableByRefColumn-- Function Call"
					strStatus = "FAIL"
					blnObjectError=True   
            				
			End If 
			
		Else
			Call ReporterFunction(strLibraryFileName,"SelectRadioButtonTableByRefColumn","1","Table Data Field","Table Object Not Found .Please Verify The Screen")
			strMsg = "Table Not Found- Please Check the Screen"
			strStatus = "FAIL"
			blnObjectError=True
		End If
	Else
		Call ReporterFunction(strLibraryFileName,"SelectRadioButtonTableByRefColumn","1","Table Data Field","Paramaters Not Passed Properly.Check the--SelectRadioButtonTableByRefColumn--Function Call")
		strMsg = "Parameters Not Passed Properly.Check the--SelectRadioButtonTableByRefColumn--Function Call"
		strStatus = "FAIL"
	End If
	If  blnObjectError  Then
		    Environment.Value("blnFatalError")=True
	   End If

	   If strStatus = "FAIL"  Then
		    SelectRadioButtonTableByRefColumn = strMsg
		    blnMainFailFlag = True
		    ImagePath=CaptureScreenshot(strStepName,objTable,False,False,False)
	   Else
		    SelectRadioButtonTableByRefColumn = strMsg
	   End If
	
	   If blnDefault_eSwiftReporting Then  
		   Call UpdateResultHtml(strStepName,refCellValue&";"&OnOffStatus,strMsg,strStatus,"")
	   End If
	   ' //////////Word Document ///////////
				If blnCreateTrainingDoc Then
					Check=CheckStatus_Doc(strStatus)
					If Check=1 Then
						Call storeCoordinates(objTable)
						Call StoreSteps(strStepName)
					End If
				End If
 ' //////////Word Document ///////////
	Set objTable=Nothing
	End if
End Function 
'************************************************************************************
'End Function - SelectCheckBoxTableByRefColumn
'************************************************************************************  

