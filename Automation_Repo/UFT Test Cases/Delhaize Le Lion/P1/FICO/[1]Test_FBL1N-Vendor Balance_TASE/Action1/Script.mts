'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_FBL1N-Vendor Balance_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_FBL1N-Vendor Balance_TASE"
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'''--------TransactionCode-FB65----------''''
'
'Call SetTcode(DT_SAPTRANSACTIONCODE)     
'Call PressEnter()     
'Call TakeScreenShot()
'Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)



'''''''--------TransactionCode-/fbl1n----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SelectRadioButton("X_AISEL", "All items", False)
Call SetTextbox("Company code","KD_BUKRS-LOW","",DT_FBL1N_1000_COMPANY_CODE,False)
Call TakeScreenShot
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FBL1N_1000_VENDOR_ACCOUNT,False)
Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call TakeScreenShot
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)
'Call FocusTextBox("Document Number","%%DYN012-LOW",False)
'Call ClickButton("SAPLSKBHTC_WRITE_LIST",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBL1N_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FBL1N_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)",true)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Change layout...   \(Ctrl\+F8\)",False)
''' SelectRowGuiGrid(gridTitle, gridIndex, columnName, refValue, blnIsItPopup)
call SelectRowGuiGrid("Displayed Columns","","Column Name","Document number",True)
Call ClickButtonIfExist("APP_FL_SING",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
'Call SelectRowGuiTable("SAPLSKBHTC_WRITE_LIST","Column content","Document number",True)
'Call ClickButtonIfExist("Hide Sel\. Fields \(Ctrl\+F2\)",True)
'Call ClickButtonIfExist("Copy   \(Enter\)",True)
''' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
'Call VerifyifGuiLabelExistsByRelativeid(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT,"wnd\[0\]/usr/lbl\[102,10\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT,"wnd\[0\]/usr/lbl\[102,11\]")
Call SelectMenuBar("Edit;Select All")
Call TakeScreenShot
Call ClickButton("btn\[45\]",False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call TakeScreenShot
Call SetTextbox("Text","\*BSEG-SGTXT","",DT_FBL1N_0100_TEXT,True)
Call TakeScreenShot
Call ClickButton("Execute changes   \(Enter\)",True)
Call VerifyStatusBar(DT_FBL1N_0500_CHECK_TEXT_OF_STATUSBAR)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Change layout...   \(Ctrl\+F8\)",False)
call SelectRowGuiGrid("Displayed Columns","","Column Name","Document number",True)
Call ClickButtonIfExist("APP_FL_SING",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
'Call SelectRowGuiTable("SAPLSKBHTC_WRITE_LIST","Column content","Document number",True)
'Call ClickButtonIfExist("Hide Sel\. Fields \(Ctrl\+F2\)",True)
'Call ClickButtonIfExist("Copy   \(Enter\)",True)
Call TakeScreenShot
Call SetHorizontalScrollBar(130,False)
''' VerifyGridCellContentbyName(gridName, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContentbyName("shell",1,"Text","",DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1)
Call VerifyGridCellContentbyName("shell",2,"Text","",DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT_OCC1)
'Call VerifyifGuiLabelExistsByRelativeid(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1,"wnd\[0\]/usr/lbl\[135,10\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT_OCC1,"wnd\[0\]/usr/lbl\[135,11\]")
'Call ClickButtonIfExist("Change layout...   \(Ctrl\+F8\)",False)
'Call ClickButtonIfExist("Copy   \(Enter\)",True)
''Call VerifyifGuiLabelExists_ByIndex(DT_FBL1N_3010_TABLECELL_SINGLE_VALUE_0,0)
''Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT),0)
''Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR,0)
''Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FAEDT),0)
'''Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM,0)
'''Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ,0)
''Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,0)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
