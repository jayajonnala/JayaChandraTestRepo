

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.14.02 Upload GL Accounting Document with Internal Order
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

gstrTestCaseName = "Test_09.07.01.14.02 Upload GL Accounting Document with Internal Order"
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
Call SetTextbox("Session","P_SESS","",DT_ZFIGL_UPLOAD_POST_1000_SESSION,False) 
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

if Dialog("text:=Microsoft Excel").Exist(5) Then
Dialog("text:=Microsoft Excel").WinButton("text:=&No").Click
End If


Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Yes",True)
Call TakeScreenShot()

''''''--------TransactionCode-SM35----------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0120_OKCD)     
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
Call ClickButtonifExist("Go back to batch input session overview   \(Enter\)",True)
Call TakeScreenShot

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", DT_ROW_OUTPUT, False)
Call TakeScreenShot
Call ClickButton("Analyze session   \(F2\)",False)

DT_TABNAME = " Log created on "&ConvertDate(DT_ZFIGL_UPLOAD_POST_0100_LOG_CREATED_ON_DATE)
Call SelectTab("TAB_DYNPRO", DT_TABNAME, False)
Call TakeScreenShot' 

Call SelectRadioButton("RB-PRO_TCODE", "Transaction", False)
Call TakeScreenShot
'
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message","2", "", "", "DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_0_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_0_OUTPUT",DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_0)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Last transaction   \(Ctrl\+F6\)",False)
Call TakeScreenShot

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message","2", "", "", "DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_0_OCC2_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_0_OCC2_OUTPUT",DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_0_OCC2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectRadioButton("RB-PRO_SESSION","Folder",False)
Call VerifyTableCellContent(4, "Message", "RSBDC_ANALYSETC_PROTOCOL", Lcase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3))
Call VerifyTableCellContent(5, "Message", "RSBDC_ANALYSETC_PROTOCOL", Lcase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4))

''''''--------TransactionCode-FBL5N ----------''''
'
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Posting date","SO_BUDAT-LOW","",ConvertDate(DT_TODAY),False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_ZFIGL_UPLOAD_POST_1000_COMPANY_CODE,False)
Call CliCkButton("%_DD_KUNNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_ZFIGL_UPLOAD_POST_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_ZFIGL_UPLOAD_POST_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_ZFIGL_UPLOAD_POST_3010_TABLECELL_SINGLE_VALUE_2,True)
Call SetTableData("SAPLALDBSINGLE","Single value","4","","",DT_ZFIGL_UPLOAD_POST_3010_TABLECELL_SINGLE_VALUE_3,True)
Call SetTableData("SAPLALDBSINGLE","Single value","5","","",DT_ZFIGL_UPLOAD_POST_3010_TABLECELL_SINGLE_VALUE_4,True)

Call ClickButtonIfExist("Copy   \(F8\)",True)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ActivateNodeGuiTree(0, "Documents;Document Number")
Call ClickButton("%_%%DYN010_%_APP_%-VALU_PUSH",False)

Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call TakeScreenShot' 
Call SetTableData("SAPLALDBINTERVAL","Lower limit","1","","","",True)
Call SetTableData("SAPLALDBINTERVAL","Upper limit","1","","","",True)

Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ActivateNodeGuiTree(0, "Documents;Document Number")
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call TakeScreenShot' 
Call SetTableData("SAPLALDBINTERVAL","Lower limit","1","","",DT_ZFIGL_UPLOAD_POST_3020_TABLECELL_LOWER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL","Upper limit","1","","",DT_ZFIGL_UPLOAD_POST_3020_TABLECELL_UPPER_LIMIT_0,True)

Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyStatusBar(Lcase(DT_ZFIGL_UPLOAD_POST_0500_CHECK_TEXT_OF_STATUSBAR))
Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZFIGL_UPLOAD_POST_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_ZFIGL_UPLOAD_POST_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

'''The Profit center numbers for the account number 11334967 are not displaying in order to verify

'Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
'Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)
'Call VerifyGridCellContent("", 3, "PRCTR", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
'Call VerifyGridCellContent("", 6, "PRCTR", 0, DT_PROF_CNTR_5)
'Call VerifyGridCellContent("", 7, "PRCTR", 0, DT_PROF_CNTR_6)
'Call VerifyGridCellContent("", 8, "PRCTR", 0, DT_PROF_CNTR_7)
'''Call VerifyGridCellContent("", 11, "PRCTR", 0, DT_PROF_CNTR_10)
'''Call VerifyGridCellContent("", 12, "PRCTR", 0, DT_PROF_CNTR_11)
'''Call VerifyGridCellContent("", 13, "PRCTR", 0, DT_PROF_CNTR_12)
'''Call VerifyGridCellContent("", 14, "PRCTR", 0, DT_PROF_CNTR_13)
'''Call VerifyGridCellContent("", 15, "PRCTR", 0, DT_PROF_CNTR_14)
'''Call VerifyGridCellContent("", 16, "PRCTR", 0, DT_PROF_CNTR_15)
'''Call VerifyGridCellContent("", 17, "PRCTR", 0, DT_PROF_CNTR_16)
'''Call VerifyGridCellContent("", 18, "PRCTR", 0, DT_PROF_CNTR_17)
'''Call VerifyGridCellContent("", 19, "PRCTR", 0, DT_PROF_CNTR_18)
'''Call VerifyGridCellContent("", 20, "PRCTR", 0, DT_PROF_CNTR_19)
'''Call VerifyGridCellContent("", 21, "PRCTR", 0, DT_PROF_CNTR_20)
'''Call VerifyGridCellContent("", 22, "PRCTR", 0, DT_PROF_CNTR_21)
'''Call VerifyGridCellContent("", 23, "PRCTR", 0, DT_PROF_CNTR_22)
'''Call VerifyGridCellContent("", 24, "PRCTR", 0, DT_PROF_CNTR_23)
'''Call VerifyGridCellContent("", 25, "PRCTR", 0, DT_PROF_CNTR_24)
'''Call VerifyGridCellContent("", 26, "PRCTR", 0, DT_PROF_CNTR_25)
'''Call VerifyGridCellContent("", 27, "PRCTR", 0, DT_PROF_CNTR_26)
'''Call VerifyGridCellContent("", 28, "PRCTR", 0, DT_PROF_CNTR_27)
'''Call VerifyGridCellContent("", 29, "PRCTR", 0, DT_PROF_CNTR_28)
'''Call VerifyGridCellContent("", 30, "PRCTR", 0, DT_PROF_CNTR_29)
'''
'''Call VerifyGridCellContent("", 33, "PRCTR", 0, DT_PROF_CNTR_32)
'''Call VerifyGridCellContent("", 34, "PRCTR", 0, DT_PROF_CNTR_33)
'''Call VerifyGridCellContent("", 35, "PRCTR", 0, DT_PROF_CNTR_34)
'''Call VerifyGridCellContent("", 38, "PRCTR", 0, DT_PROF_CNTR_37)
'''Call VerifyGridCellContent("", 39, "PRCTR", 0, DT_PROF_CNTR_38)

Call VerifyGridCellContent("", 4, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ICO_AUGP)
Call VerifyGridCellContent("", 5, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ICO_AUGP)
Call VerifyGridCellContent("", 5, "DMSHB", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_DMSHB)
Call VerifyGridCellContent("", 9, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_ICO_AUGP)
Call VerifyGridCellContent("", 10, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_ICO_AUGP)
Call VerifyGridCellContent("", 10, "DMSHB", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_DMSHB)
Call VerifyGridCellContent("", 31, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_30_ICO_AUGP)
Call VerifyGridCellContent("", 32, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_31_ICO_AUGP)
Call VerifyGridCellContent("", 32, "DMSHB", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_31_DMSHB)
Call VerifyGridCellContent("", 36, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_35_ICO_AUGP)
Call VerifyGridCellContent("", 37, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_36_ICO_AUGP)
Call VerifyGridCellContent("", 37, "DMSHB", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_36_DMSHB)
Call VerifyGridCellContent("", 40, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_39_ICO_AUGP)
Call VerifyGridCellContent("", 41, "ICO_AUGP", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_40_ICO_AUGP)
Call VerifyGridCellContent("", 41, "DMSHB", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_40_DMSHB)
Call VerifyGridCellContent("", 42, "DMSHB", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_41_DMSHB)

''''''--------TransactionCode-FAGLL03----------''''
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_ZFIGL_UPLOAD_POST_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_ZFIGL_UPLOAD_POST_1000_COMPANY_CODE_OCC1,False)
Call SelectRadioButton("X_AISEL","All items", False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(Date),False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#4;#1")
Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)

Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call TakeScreenShot' 
Call SetTableData("SAPLALDBINTERVAL","Lower limit","1","","",DT_ZFIGL_UPLOAD_POST_3020_TABLECELL_LOWER_LIMIT_0_OCC1,True)
Call SetTableData("SAPLALDBINTERVAL","Upper limit","1","","",DT_ZFIGL_UPLOAD_POST_3020_TABLECELL_UPPER_LIMIT_0_OCC1,True)

Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call ClickButtonIfExist("Execute   \(F8\)",False)

Call VerifyStatusBar(Lcase(DT_ZFIGL_UPLOAD_POST_0500_CHECK_TEXT_OF_STATUSBAR_OCC1))
Call VerifyGridCellContent("", 32, "DMSHB", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_31_DMSHB_OCC1)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZFIGL_UPLOAD_POST_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_ZFIGL_UPLOAD_POST_0841_SEARCH_DIRCT_OCC1)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "BLART", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "MWSKZ", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR_OCC1)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_PROFANDCOSTCNTR_1)
Call VerifyGridCellContent("", 3, "PRCTR", 0, DT_PROFANDCOSTCNTR_2)
Call VerifyGridCellContent("", 4, "PRCTR", 0, DT_PROFANDCOSTCNTR_3)
Call VerifyGridCellContent("", 5, "PRCTR", 0, DT_PROFANDCOSTCNTR_4)
Call VerifyGridCellContent("", 6, "PRCTR", 0, DT_PROFANDCOSTCNTR_5)
Call VerifyGridCellContent("", 7, "PRCTR", 0, DT_PROFANDCOSTCNTR_6)
Call VerifyGridCellContent("", 8, "PRCTR", 0, DT_PROFANDCOSTCNTR_7)
Call VerifyGridCellContent("", 9, "PRCTR", 0, DT_PROFANDCOSTCNTR_8)
Call VerifyGridCellContent("", 10, "PRCTR", 0, DT_PROFANDCOSTCNTR_9)
Call VerifyGridCellContent("", 11, "PRCTR", 0, DT_PROFANDCOSTCNTR_10)
Call VerifyGridCellContent("", 12, "PRCTR", 0, DT_PROFANDCOSTCNTR_11)
Call VerifyGridCellContent("", 13, "PRCTR", 0, DT_PROFANDCOSTCNTR_12)
Call VerifyGridCellContent("", 14, "PRCTR", 0, DT_PROFANDCOSTCNTR_13)
Call VerifyGridCellContent("", 15, "PRCTR", 0, DT_PROFANDCOSTCNTR_14)
Call VerifyGridCellContent("", 16, "PRCTR", 0, DT_PROFANDCOSTCNTR_15)
Call VerifyGridCellContent("", 17, "PRCTR", 0, DT_PROFANDCOSTCNTR_16)
Call VerifyGridCellContent("", 18, "PRCTR", 0, DT_PROFANDCOSTCNTR_17)
Call VerifyGridCellContent("", 19, "PRCTR", 0, DT_PROFANDCOSTCNTR_18)
Call VerifyGridCellContent("", 20, "PRCTR", 0, DT_PROFANDCOSTCNTR_19)
Call VerifyGridCellContent("", 21, "PRCTR", 0, DT_PROFANDCOSTCNTR_20)
Call VerifyGridCellContent("", 22, "PRCTR", 0, DT_PROFANDCOSTCNTR_21)
Call VerifyGridCellContent("", 23, "PRCTR", 0, DT_PROFANDCOSTCNTR_22)
Call VerifyGridCellContent("", 24, "PRCTR", 0, DT_PROFANDCOSTCNTR_23)
Call VerifyGridCellContent("", 25, "PRCTR", 0, DT_PROFANDCOSTCNTR_24)
Call VerifyGridCellContent("", 26, "PRCTR", 0, DT_PROFANDCOSTCNTR_25)

Call LogOff()
Call FinalStatus()

