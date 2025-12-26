
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Uploading the forecasted data to SAP
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

'''below two line code is added to create a run time folder for the results storing of this scenario and updating it to the result folder path.
'gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)
gstrTestCaseName = "TC_01_Test_Uploading the forecasted data to SAP_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''

Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)
Call ExcelFileTransfer(DT_ForeCastSourceFolderPath,DT_ForeCastTargetFilePath,DT_ForeCastTargetFile,36)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''----------------------Tcode ZMDPP_ARGOS_FORECAST ----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()


Call SetTextbox("Site","P_WERKS","",DT_SITE,False)
Call SetTextbox("Requirements type","P_BEDAE","",DT_REQ_TYPE,False)
Call SetTextbox("Version","P_VERSB","",DT_VERSION,False)
Call SetTextbox("Date type","P_ENTLI","",DT_DATATYPE,False)

Call TakeScreenShot()
Call SelectCheckbox("P_SKIP1",0,DT_SKIPFIRSTROW,False)
Call SelectCheckbox("P_ABCUP",0,DT_ABCUPDATE,False)

'Call SetTextbox("File name","P_FILNM","",DT_UPLOAD_EXCEL_FILEPATH,False)
Call SetTextbox("File name","P_FILNM","",DT_ForeCastTargetFilePath,False)
Call TakeScreenshot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call SelectColumnGuiGrid("", 0, "Quantity", False)
Call TakeScreenShot()
Call ClickButton("Sort in descending order   \(Ctrl\+Shift\+F4\)",False)
Call TakeScreenShot()

Call SelectColumnGuiGrid("", 0, "Exception", False)
Call TakeScreenShot()
Call ClickButton("Set filter   \(Ctrl\+F5\)",False)
Call TakeScreenShot()

Call SelectCheckbox("P_YELLOW", 0, "ON", True)
Call SelectCheckbox("P_GREEN",0, "ON", True)
Call SelectCheckbox("P_LIGHT", 0, "ON", True)
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot()

Call GetGridContentByTitle("", "", "Material", 1, "DT_Material")
Call GetGridContentByTitle("", "", "Quantity", 1, "DT_QUNAITY")
Call GetGridContentByTitle("", "", "Date", 1, "DT_Date")

Call GetGridContentByTitle("", "", "Material", 2, "DT_Material_OCC1")
Call GetGridContentByTitle("", "", "Quantity", 2, "DT_QUNAITY_OCC1")
Call GetGridContentByTitle("", "", "Date", 2, "DT_Date_OCC1")

Call GetGridContentByTitle("", "", "Material", 3, "DT_Material_OCC2")
Call GetGridContentByTitle("", "", "Quantity", 3, "DT_QUNAITY_OCC2")
Call GetGridContentByTitle("", "", "Date", 3, "DT_Date_OCC2")

Call SelectRowRangeGuiGrid("", "", 1, 3, False)
Call TakeScreenShot()
Call ClickBUtton("Create Requirements   \(F8\)",False)
Call TakeScreenShot()

Call VerifyGridCellContentbyName("shell", 1,  "Requirements Status", "", DT_VERIFY_REQUIREMENT_STATUS)
Call VerifyGridCellContentbyName("shell", 2,  "Requirements Status", "", DT_VERIFY_REQUIREMENT_STATUS)
Call VerifyGridCellContentbyName("shell", 3,  "Requirements Status", "", DT_VERIFY_REQUIREMENT_STATUS)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

'''----------------------Tcode: MD63 ----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_SAPTRANSACTIONCODE_OCC1)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call Selectradiobutton("AM60X-MATAW","Article",False)
Call SetTextbox("Article","AM60X-MATNR", "", DT_Material, False)
Call Selectradiobutton("AM60X-VERAK","All active versions",False)
Call SetTextbox("Site","AM60X-WERKS", "", DT_SITE, False)
Call SetTextbox("Selected version","RM60X-VERSB", "", DT_SELECTED_VERSION, False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call VerifyTableCellContent(1, DT_WEEK, "SAPLM60ETC_SHEET", Replace(DT_QUNAITY,",","."))
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call Selectradiobutton("AM60X-MATAW","Article",False)
Call SetTextbox("Article","AM60X-MATNR", "", DT_Material_OCC1, False)
Call Selectradiobutton("AM60X-VERAK","All active versions",False)
Call SetTextbox("Site","AM60X-WERKS", "", DT_SITE, False)
Call SetTextbox("Selected version","RM60X-VERSB", "", DT_SELECTED_VERSION, False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call VerifyTableCellContent(1, DT_WEEK_OCC1, "SAPLM60ETC_SHEET", Replace(DT_QUNAITY_OCC1,",","."))
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call Selectradiobutton("AM60X-MATAW","Article",False)
Call SetTextbox("Article","AM60X-MATNR", "", DT_Material_OCC2, False)
Call Selectradiobutton("AM60X-VERAK","All active versions",False)
Call SetTextbox("Site","AM60X-WERKS", "", DT_SITE, False)
Call SetTextbox("Selected version","RM60X-VERSB", "", DT_SELECTED_VERSION, False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call VerifyTableCellContent(1, DT_WEEK_OCC2, "SAPLM60ETC_SHEET", Replace(DT_QUNAITY_OCC2,",","."))

Call LogOff()
Call FinalStatus()
