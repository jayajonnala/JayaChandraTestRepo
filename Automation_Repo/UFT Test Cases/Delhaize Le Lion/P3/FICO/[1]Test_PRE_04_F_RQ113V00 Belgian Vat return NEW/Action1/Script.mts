'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_04_F_RQ113V00 Belgian Vat return NEW 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_04_F_RQ113V00 Belgian Vat return NEW"
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-ZFIGL_UPLOAD_POST----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call TakeScreenShot

''''''--------TransactionCode--SM35---------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

'Call SetTextbox("Created by:","D0100-CREATOR","",DT_ZFIGL_UPLOAD_POST_1005_CREATED_BY,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1, False)
Call ClickButton("Process session   \(F8\)",False)
Call SelectRadioButton("D0300-ERROR", "Display errors only", True)
Call TakeScreenShot
Call ClickButton("Process   \(Enter\)",True)
Call TakeScreenShot
wait 60
Call ClickButton("Go back to batch input session overview   \(Enter\)",True)
wait 60
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1, False)
Call ClickButton("Analyze session   \(F2\)",False)
wait 60
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1, False)
Call ClickButton("Analyze session   \(F2\)",False)
wait 60
'Call VerifyTextBoxContent("Processed", "APQDCNT-TRANSCNTF", "", DT_ZFIGL_UPLOAD_POST_0200_CHECK_TEXT_OF_APQDCNTTRANSCNTF, False)
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_SM35_1005_TO__OCC1),False)
Call SelectRadioButton("RB-PRO_TCODE", "Transaction", False)
Call TakeScreenShot
' GetTableCellData(tableName, columnName, rowNumber, refColumnName, refCellValue, dataTableColumnName, blnIsItPopup)

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message","", "No.", "312","DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2", False)
Call SelectRadioButton("RB-PRO_SESSION", "Folder", False)
Call TakeScreenShot
Call VerifyTableCellContent(5, "Message", "RSBDC_ANALYSETC_PROTOCOL", Lcase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4))
Call VerifyTableCellContent(4, "Message", "RSBDC_ANALYSETC_PROTOCOL", LCase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3))
'''GetRowNo =2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
