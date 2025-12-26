		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.13.1Store Cash Management_Electricity (DEH) manual p
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

gstrTestCaseName = "Test_09.07.01.02.13.1Store Cash Management_Electricity (DEH) manual p"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-ZFIGL_UPLOAD_POST----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetSpecialTextbox("File path name","P_FILE", "",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME, False)
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

If Dialog("text:=Microsoft Excel").Exist(5) Then
	Dialog("text:=Microsoft Excel").WinButton("text:=&No").Click
End If

Call ClickButton("Yes",True)
Call TakeScreenShot

''''''--------TransactionCode-SM35----------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call TakeScreenShot
Call ClickButton("Process session   \(F8\)",False)
Call TakeScreenShot
Call SelectRadioButton("D0300-ERROR", "Display errors only", True)
Call TakeScreenShot
Call ClickButtonifExist("Process   \(Enter\)",True)
Call TakeScreenShot
Wait 5

Call ClickButtonifExist("Go back to batch input session overview   \(Enter\)",True)
Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)	
Call TakeScreenShot
Call ClickButton("Analyze session   \(F2\)",False)
		
DT_TABNAME = " Log created on "&ConvertDate(DT_ZFIGL_UPLOAD_POST_0100_LOG_CREATED_ON_DATE)
Call SelectTab("TAB_DYNPRO",DT_TABNAME,False)
Call TakeScreenShot
Call SelectRadioButton("RB-PRO_TCODE", "Transaction",False)
Call TakeScreenShot
Wait 5
Call SelectCheckbox("RB-LOG_DETAIL", 0, "ON", False)
Wait 2
Call SelectCheckbox("RB-LOG_DETAIL", 0, "OFF", False)

Call FindRowNumber("RSBDC_ANALYSETC_PROTOCOL", "No.", "312", "DT_ROW_OUTPUT")

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",DT_ROW_OUTPUT, "", "", "DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3_OUTPUT",DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)




''''''--------TransactionCode-FBL1N ----------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_ZFIGL_UPLOAD_POST_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_ZFIGL_UPLOAD_POST_3010_TABLECELL_SINGLE_VALUE_0,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZFIGL_UPLOAD_POST_0500_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_ZFIGL_UPLOAD_POST_0500_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
'
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call LogOff'

