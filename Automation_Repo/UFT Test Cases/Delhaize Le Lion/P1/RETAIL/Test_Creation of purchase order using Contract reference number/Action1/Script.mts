
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Creation of purchase order using Contract reference number_TASE
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


gstrTestCaseName = "TC_02_Test_Creation of purchase order using Contract reference number_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------Getting Contract Reference Nummber----------------'''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)

Call ExcelFileTransfer(DT_ContractSourceFolderPath,DT_ContractTargetFilePath,DT_ContractTargetFile,31)

Call GetExcelColumnNumber(DT_ContractTargetFilePath,DT_SHEET_NAME,DT_COLUMN_NAME,DT_ROWNUMBER,"DT_CONTRACT_STATUS_COLUMN_NUM")
Call GetExcelColumnNumber(DT_ContractTargetFilePath,DT_SHEET_NAME,DT_COLUMN_NAME_OCC1,DT_ROWNUMBER,"DT_OPEN_CONTRACT_QUNAITY_COLUMN_NUM")
Call GetExcelColumnNumber(DT_ContractTargetFilePath,DT_SHEET_NAME,DT_COLUMN_NAME_OCC2,DT_ROWNUMBER,"DT_CONTARCT_REFERENCE_COLUMN_NUM")

Call GetExcelCellValueWithReference(DT_ContractTargetFilePath,DT_SHEET_NAME,"Contract Reference","Is Active","1000",3,DT_CONTRACT_STATUS_COLUMN_NUM,DT_OPEN_CONTRACT_QUNAITY_COLUMN_NUM,DT_CONTACT_REFERENCE_COLUMN_NUM,"DT_CONTRACT_REFERENCE_NUMBER")

'''''''--------------login----------------'''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''----------------------Tcode ME21N----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetCombo("MEPO_TOPLINE-BSART", DT_COMBO_VALUE)
Call TakeScreenShot()

Call ClickButtonIfexist("Expand Items Ctrl\+F3",False)
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211", "Outline agreement", 1, "", "", DT_CONTRACT_NUMBER, False)
Call PressEnter()
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211", "PO Quantity", 1, "", "", DT_PO_QUANTITY, False)
Call PressEnter()
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211", "PO Quantity", 1, "", "", DT_PO_QUANTITY, False)
Call PressEnter()
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211", "Deliv. Date", 1, "", "", ConvertDate(DT_DELIVERY_DATE), False)
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Messages   \(Shift\+F9\)",False)
Call TakeScreenShot()

Call GetNumberOfRows("SAPDV70ATC_NAST3", 0, "DT_NUMOFROWS")

For i = 1 To DT_NUMOFROWS
		Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
		Call FindRowNumber("SAPDV70ATC_NAST3", "Status", "S_TL_Y", "DT_ROW_NUMBER")
	If DT_ROW_NUMBER <> Empty Then
			Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
			CAll SelectRowGuiTableByRow("SAPDV70ATC_NAST3", DT_ROW_NUMBER, False)
			Call TakeScreenShot()
			Call ClickBUtton("Delete \(once\)   \(Shift\+F6\)",False)
			Call TakeScreenShot()
			Call ClickButtonifExist("Continue   \(Enter\)",False)
			Call TakeScreenShot()
			Wait 5
			i = i-1
			Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
			Call GetNumberOfRows("SAPDV70ATC_NAST3", 0, "DT_NUMOFROWS")	
			Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
			Else
			Exit For
	End If
	
Next

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call ClickButton("Check   \(Ctrl\+Shift\+F3\)",False)
Call TakeScreenShot()

Call VerifyStatusBar(DT_STATUSBAR_MESSAGE)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call GetStatusBar("item2", "DT_PO_NUMBER_OUTPUT")
Call VerifyStatusBar(Lcase("DG Winery Base Wine created under the number "&DT_PO_NUMBER_OUTPUT&""))

Call LogOff()
Call FinalStatus()





















Public Function GetExcelCellValueWithReference(InputExcelFilePath,InputSheetName,columnName,ContractStatus,OpenContractQuantity,RowNumberValue,ContractStatusColumnNumber,OpenContractColumnNumber,ContractReferenceColumnNumber,dataTableColumnName)
	If Not (Environment.Value("blnFatalError")) Then
	   If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : GetExcelCellValueWithReference"
	strStepName = "Get Excel Cell Vale"
	               Set StrExcel=createobject("excel.application")
			 StrExcel.Visible=False

			Set StrWorkBook = StrExcel.Workbooks.Open(InputExcelFilePath)
			Set StrSheet = StrExcel.Worksheets.Item(InputSheetName&" ")

				Numofrows = StrSheet.UsedRange.Rows.Count
				Numofcolumns = StrSheet.UsedRange.Columns.Count

				For RowNumber = RowNumberValue To Numofrows
					if StrSheet.cells(RowNumber,Cint(ContractStatusColumnNumber)).value = ContractStatus Then
						If StrSheet.cells(RowNumber,Cint(OpenContractColumnNumber)).value > Cint(OpenContractQuantity) Then
							ContractReferenceNumber = StrSheet.cells(RowNumber,Cint(ContractReferenceColumnNumber)).value
							strStatus = "DONE"
							Call ReporterFunction(strLibraryFileName,"GetExcelColumnNumber","2","Gui Grid","Excel value of the contract reference number Fetched Into Datatable")
							strMsg = "Excel cell value of contract reference number  "&ContractReferenceNumber&" Fetched Into Datatable"
				Exit For
						End if
					End If
				Next
	          
			If  blnObjectError  Then
		    Environment.Value("blnFatalError")=True
	        End if

	   If strStatus = "FAIL"  Then
		    GetExcelCellValueWithReference = strMsg
		    blnMainFailFlag = True
		    ImagePath=CaptureScreenshot(strStepName,objGrid,False,False,False)
	   Else
		    GetExcelCellValueWithReference = strMsg
	   End If
	   If blnWriteDataToOutputSheet Then
	                                    strStepName = "Retrieve Contract Reference Number value:'"&ContractReferenceNumber&"'  from the excel sheet of "&columnName&" and store in '"&gstrOutputSheetName&"' sheet under column "&dataTableColumnName
									'stepName="Retrieve '"&strCapturedText&"' value from page "&browserName&" and store in "&gstrOutputSheetName&" sheet under column Webelement_"&webElementClassORhtmlId
										call WriteRunTimeDataToExcel (dataTableColumnName,ContractReferenceNumber)
										'call WriteRunTimeDataToExcel ("ID_"&webElementClassORhtmlId,strCapturedText)
							Else
										strStepName = "Retrieve '"&ContractReferenceNumber&"' value from the grid"
										'stepName="Retrieve '"&strCapturedText&"' value from page "&browserName
		End If 
	   If blnDefault_eSwiftReporting Then  
		   Call UpdateResultHtml(strStepName,ContractReferenceNumber,strMsg,strStatus,"")
	   End If
	End If
	
	StrExcel.Quit
Set StrSheet=Nothing
Set StrWorkBook=Nothing
Set StrExcel=Nothing

End Function

