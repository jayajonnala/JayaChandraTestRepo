'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.04 Store Cash Management_Ticket restaurant_P1
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

gstrTestCaseName = "Test_09.07.01.02.04 Store Cash Management_Ticket restaurant_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-ZFIGL_UPLOAD_POST----------''''
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False) 
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Yes",True)
Call TakeScreenShot()

''''''--------TransactionCode-SM35----------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call FindRowNumber("SAPMSBDC_CCTC_APQI", "Created By", "AUTOFAB", "DT_ROW_OUTPUT")

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", DT_ROW_OUTPUT, False)
Call TakeScreenShot
Call ClickButton("Process session   \(F8\)",False)
Call TakeScreenShot
Call SelectRadioButton("D0300-ERROR", "Display errors only", True)
Call TakeScreenShot
Call ClickButtonifExist("Process   \(Enter\)",True)
Call TakeScreenShot

''These lines of code are not required, and it is implemented as temporary workaround
''Call SetTextbox("Posting Date","BKPF-BUDAT", 0, ConvertDate(DT_ZFIGL_UPLOAD_POST_0100_DOCUMENT_DATE), False)
''Call SetTextbox("Period","BKPF-MONAT", 0, ConvertDoubleDigit(Cstr(Month(DT_ZFIGL_UPLOAD_POST_0100_DOCUMENT_DATE))), False)
''Call TakeScreenShot
''Call PressEnter()
''Call takeScreenShot()

Call ClickButtonifExist("Go back to batch input session overview   \(Enter\)",True)
Call TakeScreenShot
Wait 60

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", DT_ROW_OUTPUT, False)
Call TakeScreenShot
Wait 60
Call ClickButton("Analyze session   \(F2\)",False)
Call SelectTab("TAB_DYNPRO", " Dynpros", False)
Call TakeScreenShot'
DT_TABNAME = " Log created on "&ConvertDate(DT_ZFIGL_UPLOAD_POST_1005_FROM)
Call SelectTab("TAB_DYNPRO", DT_TABNAME, False)
Call TakeScreenShot' 

Call FindRowNumber("RSBDC_ANALYSETC_PROTOCOL", "No.", "312", "DT_ROW_OUTPUT")
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",DT_ROW_OUTPUT, "", "", "DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7_OUTPUT",DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''''''--------TransactionCode-FAGLL03----------''''
''
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#3;#1")

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ZFIGL_UPLOAD_POST_0100_DOCUMENT_NUMBER,False)
Call PressEnter()
Call ClickBUtton("Back   \(F3\)",False)
Call ClickBUtton("Yes",True)
Wait 2
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot  

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZFIGL_UPLOAD_POST_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)


Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 3, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
Call VerifyGridCellContent("", 4, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLART)
Call VerifyGridCellContent("", 7, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BLART)
Call VerifyGridCellContent("", 8, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_BLART)
Call VerifyGridCellContent("", 9, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_BLART)
Call VerifyGridCellContent("", 10, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_BLART)
'''Call VerifyGridCellContent("", 13, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_12_BLART)
'''Call VerifyGridCellContent("", 14, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_13_BLART)
'''Call VerifyGridCellContent("", 15, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_14_BLART)
'''Call VerifyGridCellContent("", 16, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_14_BLART)
Call VerifyGridCellContent("", 13, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_BLART)
Call VerifyGridCellContent("", 14, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_BLART)
Call VerifyGridCellContent("", 15, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_BLART)
Call VerifyGridCellContent("", 16, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_BLART)

Call VerifyGridCellContent("", 1, "MWSKZ", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)

Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 3, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
Call VerifyGridCellContent("", 4, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HKONT)
Call VerifyGridCellContent("", 7, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_HKONT)
Call VerifyGridCellContent("", 8, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_HKONT)
Call VerifyGridCellContent("", 9, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_HKONT)
Call VerifyGridCellContent("", 10, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_HKONT)
'''Call VerifyGridCellContent("", 13, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_12_HKONT)
'''Call VerifyGridCellContent("", 14, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_13_HKONT)
'''Call VerifyGridCellContent("", 15, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_14_HKONT)
'''Call VerifyGridCellContent("", 16, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_14_HKONT)
Call VerifyGridCellContent("", 13, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_HKONT)
Call VerifyGridCellContent("", 14, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_HKONT)
Call VerifyGridCellContent("", 15, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_HKONT)
Call VerifyGridCellContent("", 16, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_HKONT)

Call LogOff'
Call FinalStatus()

